using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Discord;
using Discord.WebSocket;
using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Configuration;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System.ComponentModel.DataAnnotations.Schema;

namespace ROCAPointBot
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);
            builder.Services.AddHostedService<DiscordBotService>();
            var app = builder.Build();
            app.MapGet("/", () => "✅ ROCA Point Bot 正在運行中！(包含自動同步成員功能)");
            app.Run();
        }

        public static DateTime GetTaipeiTime()
        {
            DateTime utcNow = DateTime.UtcNow;
            try { return TimeZoneInfo.ConvertTimeFromUtc(utcNow, TimeZoneInfo.FindSystemTimeZoneById("Taipei Standard Time")); }
            catch { return utcNow.AddHours(8); }
        }
    }

    public class DiscordBotService : BackgroundService
    {
        private DiscordSocketClient _client;
        private static readonly HttpClient _http = new HttpClient();
        private readonly string _discordToken;
        private readonly IConfiguration _configuration;

        // 👇 新增這個變數：用來記錄機器人是否已經完成初次開機
        private bool _isReady = false;
        // 👇 新增：你的開發者 ID 與私訊頻道 ID
        private readonly ulong _developerId = 1018018187318673471;
        private readonly ulong _devDmChannelId = 1476489806946242592;
        // 👇 新增這兩個靜態變數，用來暫存追蹤中的語音頻道與結算面板
        private static Dictionary<ulong, VoiceEventSession> _activeEvents = new();
        private static Dictionary<string, PendingEventDistribution> _pendingDistributions = new();
        public DiscordBotService(IConfiguration config)
        {
            _configuration = config;
            _discordToken = config["DiscordToken"];
        }

        // 👇 新增：傳送狀態私訊給開發者的功能
        private async Task SendStatusToDeveloperAsync(string message)
        {
            try
            {
                // 優先嘗試傳送至指定的 DM 頻道
                if (_client.GetChannel(_devDmChannelId) is IMessageChannel dmChannel)
                {
                    await dmChannel.SendMessageAsync(message);
                }
                else
                {
                    // 備用方案：直接透過 User ID 傳送
                    var dev = await _client.GetUserAsync(_developerId);
                    if (dev != null) await dev.SendMessageAsync(message);
                }
            }
            catch { /* 若沒有權限或頻道錯誤則忽略，避免機器人崩潰 */ }
        }
        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            // 在 protected override async Task ExecuteAsync 內：
            var config = new DiscordSocketConfig
            {
                GatewayIntents = GatewayIntents.AllUnprivileged | GatewayIntents.GuildMembers,
                AlwaysDownloadUsers = true // 👈 新增這行：強制機器人抓取完整伺服器名單以利計算人數
            };
            _client = new DiscordSocketClient(config);
            // 👇 1. 將這裡原本的 _client.Ready 替換成這段：
            _client.Ready += async () =>
            {
                if (_isReady)
                {
                    // 如果已經開機過，代表這是網路波動導致的「重新連線」，我們就記錄在 Console 即可，不要重複執行
                    Console.WriteLine($"🔄 [系統通知] 機器人剛剛重新連線了 ({Program.GetTaipeiTime():HH:mm:ss})");
                    return;
                }

                _isReady = true; // 標記為已成功開機

                await RegisterCommandsAsync();
                // 機器人開機時，自動發送私訊給你
                await SendStatusToDeveloperAsync($"🟢 **[系統通知]** ROCA Point Bot 已成功啟動！\n> 時間：{Program.GetTaipeiTime():yyyy-MM-dd HH:mm:ss}");
            };
            _client.SlashCommandExecuted += HandleSlashCommandAsync;
            _client.InteractionCreated += HandleInteractionAsync;
            // 👇 新增這行：註冊語音頻道進出監聽器
            _client.UserVoiceStateUpdated += HandleVoiceStateUpdatedAsync;
            await _client.LoginAsync(TokenType.Bot, _discordToken);
            await _client.StartAsync();
            // 👇 2. 新增這段：防休眠自動呼叫 (Keep-Alive) 機制
            _ = Task.Run(async () =>
            {
                while (!stoppingToken.IsCancellationRequested)
                {
                    try
                    {
                        // ⚠️ 非常重要：請將下方的網址替換成你部署網站的「真實網址」
                        string myWebsiteUrl = "http://roca-bot.runasp.net/";
                        await _http.GetAsync(myWebsiteUrl);
                        Console.WriteLine($"🟢 [Keep-Alive] 成功 Ping 網站防止休眠 ({Program.GetTaipeiTime():HH:mm:ss})");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"⚠️ [Keep-Alive] 呼叫失敗: {ex.Message}");
                    }

                    // 將間隔縮短為 2 分鐘 (原本是 5 分鐘太久了)
                    // 且將 Delay 移到請求之後，確保開機立刻執行一次
                    await Task.Delay(TimeSpan.FromMinutes(2), stoppingToken);
                }
            }, stoppingToken);
            _ = Task.Run(async () =>
            {
                try
                {
                    using var db = new BotDbContext(_configuration);
                    await db.Database.EnsureCreatedAsync();

                    // 💡 自動嘗試在現有資料表中加入新欄位 (用 Try-Catch 防止重複加入報錯)
                    try { await db.Database.ExecuteSqlRawAsync("ALTER TABLE GuildConfigs ADD ServerCode NVARCHAR(50) NULL"); } catch { }
                    // 1. 將未來新增的型別修正為 decimal(20,0)
                    try { await db.Database.ExecuteSqlRawAsync("ALTER TABLE GuildConfigs ADD MasterLogChannelId decimal(20,0) NULL"); } catch { }
                    // 2. 新增這行：強制將現有資料庫中錯誤的 BIGINT 欄位轉正為 decimal(20,0)
                    try { await db.Database.ExecuteSqlRawAsync("ALTER TABLE GuildConfigs ALTER COLUMN MasterLogChannelId decimal(20,0) NULL"); } catch { }
                    // 自動嘗試在現有資料表中加入 GoogleSheetId 欄位
                    try { await db.Database.ExecuteSqlRawAsync("ALTER TABLE GuildConfigs ADD GoogleSheetId NVARCHAR(100) NULL"); } catch { }
                    // 👇 新增這行：自動替現有的 AdminChannels 加上第二個身分組欄位
                    try { await db.Database.ExecuteSqlRawAsync("ALTER TABLE AdminChannels ADD MndAdminRoleId2 decimal(20,0) NULL"); } catch { }
                    try { await db.Database.ExecuteSqlRawAsync("ALTER TABLE GuildConfigs ADD RewardMenuContent NVARCHAR(MAX) NULL"); } catch { }
                    try { await db.Database.ExecuteSqlRawAsync("ALTER TABLE GuildConfigs ADD RewardMenuUpdateTime datetime2 NULL"); } catch { }
                    // 為了相容 SQLite 的備用語法
                    try { await db.Database.ExecuteSqlRawAsync("ALTER TABLE GuildConfigs ADD RewardMenuContent TEXT NULL"); } catch { }
                    try { await db.Database.ExecuteSqlRawAsync("ALTER TABLE GuildConfigs ADD RewardMenuUpdateTime DATETIME NULL"); } catch { }
                    // 👇 新增這兩行：自動為現有資料庫加入職位欄位
                    try { await db.Database.ExecuteSqlRawAsync("ALTER TABLE UserPoints ADD RobloxRank NVARCHAR(MAX) NULL"); } catch { }
                    try { await db.Database.ExecuteSqlRawAsync("ALTER TABLE UserPoints ADD RobloxRank TEXT NULL"); } catch { }
                    // 自動建立草稿品項資料表
                    try
                    {
                        await db.Database.ExecuteSqlRawAsync(@"
                        IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'RewardMenuItems')
                        BEGIN
                            CREATE TABLE RewardMenuItems (
                                Id int IDENTITY(1,1) PRIMARY KEY,
                                GuildId decimal(20,0) NOT NULL,
                                ItemName nvarchar(max) NOT NULL,
                                Points int NOT NULL,
                                Note nvarchar(max) NULL
                            )
                        END");
                    }
                    catch { }
                    // 相容 SQLite
                    try { await db.Database.ExecuteSqlRawAsync("CREATE TABLE IF NOT EXISTS RewardMenuItems (Id INTEGER PRIMARY KEY AUTOINCREMENT, GuildId INTEGER NOT NULL, ItemName TEXT NOT NULL, Points INTEGER NOT NULL, Note TEXT NULL)"); } catch { }
                    // 👇 ================= 新增這段程式碼 ================= 👇
                    // 自動嘗試建立 AdminChannels 資料表 (如果它不存在的話)
                    try
                    {
                        await db.Database.ExecuteSqlRawAsync(@"
                            IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'AdminChannels')
                            BEGIN
                                CREATE TABLE AdminChannels (
                                    ChannelId decimal(20,0) NOT NULL,
                                    GuildId decimal(20,0) NOT NULL,
                                    TargetServerCodes nvarchar(max) NOT NULL,
                                    CONSTRAINT PK_AdminChannels PRIMARY KEY (ChannelId)
                                )
                            END");
                    }
                    catch (Exception ex) { Console.WriteLine($"⚠️ 建立 AdminChannels 資料表失敗: {ex.Message}"); }
                    // 👆 ================================================== 👆
                    // 👇 新增這行：自動替現有的 AdminChannels 加上身分組欄位
                    try { await db.Database.ExecuteSqlRawAsync("ALTER TABLE AdminChannels ADD MndAdminRoleId decimal(20,0) NULL"); } catch { }
                    Console.WriteLine("✅ 資料庫連線成功並已準備就緒！");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ 資料庫連線失敗: {ex.Message}");
                }
            }, stoppingToken);

            await Task.Delay(Timeout.Infinite, stoppingToken);
        }

        private async Task RegisterCommandsAsync()
        {
            var commands = new List<ApplicationCommandProperties>
            {
                new SlashCommandBuilder().WithName("setup-roca").WithDescription("🔗 綁定 Roblox 群組並自動同步成員")
                .AddOption(new SlashCommandOptionBuilder()
                    .WithName("roblox_group_id")
                    .WithDescription("選擇綁定單位")
                    .WithType(ApplicationCommandOptionType.String)
                    .WithRequired(true)
                    .AddChoice("憲兵", "13549943")
                    .AddChoice("裝甲", "13662982")
                    .AddChoice("航特", "16223475"))
                // 👇 擴充至 7 個管理身分組選項
                .AddOption("admin_role_1", ApplicationCommandOptionType.Role, "主要管理身分組 (刪除資料或異常時僅會 Ping 此身分組)", isRequired: true)
                .AddOption("admin_role_2", ApplicationCommandOptionType.Role, "管理身分組 2 (選填，擁有指令查詢權限)", isRequired: false)
                .AddOption("admin_role_3", ApplicationCommandOptionType.Role, "管理身分組 3 (選填，擁有指令查詢權限)", isRequired: false)
                .AddOption("admin_role_4", ApplicationCommandOptionType.Role, "管理身分組 4 (選填)", isRequired: false)
                .AddOption("admin_role_5", ApplicationCommandOptionType.Role, "管理身分組 5 (選填)", isRequired: false)
                .AddOption("admin_role_6", ApplicationCommandOptionType.Role, "管理身分組 6 (選填)", isRequired: false)
                .AddOption("admin_role_7", ApplicationCommandOptionType.Role, "管理身分組 7 (選填)", isRequired: false)
                .Build(),
                new SlashCommandBuilder().WithName("bind-sheet")
                .WithDescription("📊 獲取本單位專屬的 Google 試算表連結與同步選單")
                .Build(),
                new SlashCommandBuilder().WithName("sync-members").WithDescription("🔄 手動同步 Roblox 群組的最新成員至資料庫").Build(),
                new SlashCommandBuilder().WithName("points").WithDescription("📊 查詢點數").AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("addpoint").WithDescription("➕ 發放點數").AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true).AddOption("points", ApplicationCommandOptionType.Integer, "點數數量", isRequired: true).AddOption("reason", ApplicationCommandOptionType.String, "原因備註", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("history").WithDescription("📜 查詢玩家近期紀錄").AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("viewall").WithDescription("🏆 點數總排行榜 (顯示所有成員)").Build(),
                new SlashCommandBuilder().WithName("del-record")
                    .WithDescription("🚨 刪除單筆紀錄")
                    .AddOption("id", ApplicationCommandOptionType.Integer, "紀錄 ID", isRequired: true)
                    .AddOption("reason", ApplicationCommandOptionType.String, "撤銷原因", isRequired: true) // 👈 新增這行原因必填
                    .Build(),
                new SlashCommandBuilder().WithName("edit-admin")
                .WithDescription("👮 新增或移除機器人管理員身分組 (需雙重確認)")
                .AddOption(new SlashCommandOptionBuilder()
                    .WithName("action")
                    .WithDescription("選擇動作")
                    .WithType(ApplicationCommandOptionType.String)
                    .WithRequired(true)
                    .AddChoice("➕ 新增身分組", "add")
                    .AddChoice("➖ 移除身分組", "remove"))
                .AddOption("role", ApplicationCommandOptionType.Role, "選擇目標身分組", isRequired: true)
                .Build(),
                new SlashCommandBuilder().WithName("unbind-roca").WithDescription("🔓 解除綁定設定").Build(),
                new SlashCommandBuilder().WithName("clear-all-data").WithDescription("🔥 【極度危險】刪除本伺服器所有資料").Build(),
                new SlashCommandBuilder().WithName("daily-records").WithDescription("📅 查詢指定日期的所有點數紀錄").AddOption("date", ApplicationCommandOptionType.String, "格式: YYYY-MM-DD", isRequired: true).Build(),
                // 👇 新增這行：註冊 status 指令
                new SlashCommandBuilder().WithName("status").WithDescription("🟢 檢查機器人目前連線與運作狀態").Build(),
                // 👇 新增這兩個指令
                 new SlashCommandBuilder().WithName("admin-setup")
                .WithDescription("🔒 國防部專用：設定監控頻道與授權身分組")
                .AddOption("server_codes", ApplicationCommandOptionType.String, "要監控的伺服器編號(多個用逗號分隔，如 A1B2,C3D4)", isRequired: true)
                .AddOption("mnd_role", ApplicationCommandOptionType.Role, "設定可以使用 /admin-view 的長官身分組 1", isRequired: true)
                // 👇 新增這行選項
                .AddOption("mnd_role_2", ApplicationCommandOptionType.Role, "設定可以使用 /admin-view 的長官身分組 2 (選填)", isRequired: false)
                .Build(),
                new SlashCommandBuilder().WithName("admin-view").WithDescription("👁️ 國防部專用：查看特定單位的總排行榜").AddOption("server_code", ApplicationCommandOptionType.String, "目標伺服器編號", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("my-code").WithDescription("🔑 查詢本伺服器的專屬資料庫編號 (交給國防部綁定通知用)").Build(),
                new SlashCommandBuilder().WithName("removepoint")
                    .WithDescription("➖ 扣除或兌換點數")
                    .AddOption("user", ApplicationCommandOptionType.User, "選擇人員", isRequired: true)
                    .AddOption("points", ApplicationCommandOptionType.Integer, "扣除數量", isRequired: true)
                    .AddOption("reason", ApplicationCommandOptionType.String, "原因備註", isRequired: true)
                    .Build(),
                // 👇 新增這行：註冊 log-channel 指令
                new SlashCommandBuilder().WithName("log-channel")
                    .WithDescription("📝 設定本伺服器的專屬紀錄(Log)推播頻道")
                    .AddOption("channel", ApplicationCommandOptionType.Channel, "選擇要接收推播的文字頻道", isRequired: true)
                    .Build(),
              new SlashCommandBuilder().WithName("add-menu-item")
                    .WithDescription("➕ 新增兌換 MENU 品項 (僅暫存於草稿，需透過 /edit-menu 發布)")
                    .AddOption("item_name", ApplicationCommandOptionType.String, "品項名稱 (若要新增分類標題，請將點數設為 0)", isRequired: true)
                    .AddOption("points", ApplicationCommandOptionType.Integer, "所需點數 (設為 0 將視為分類標題)", isRequired: true)
                    .AddOption(new SlashCommandOptionBuilder()
                        .WithName("note_type")
                        .WithDescription("選擇備註的標籤類型 (預設為「備註」)")
                        .WithType(ApplicationCommandOptionType.String)
                        .WithRequired(false)
                        .AddChoice("🏷️ 條件", "條件")
                        .AddChoice("📝 備註", "備註"))
                    .AddOption("note", ApplicationCommandOptionType.String, "內容 (選填)", isRequired: false)
                    // 👇 加上這一行：讓管理員可以選填「要插在哪個品項後面」
                    .AddOption("insert_after", ApplicationCommandOptionType.String, "要插在哪個品項的【下方】？請輸入該品項完整名稱(選填)", isRequired: false)
                    .Build(),
                new SlashCommandBuilder().WithName("remove-menu-item")
                    .WithDescription("➖ 移除草稿中的 MENU 品項")
                    .AddOption("item_name", ApplicationCommandOptionType.String, "請輸入要移除的完整品項名稱", isRequired: true).Build(),

                new SlashCommandBuilder().WithName("change-menu-item")
                    .WithDescription("🔄 替換草稿中的 MENU 品項 (直接覆寫舊品項以保持原本的排列順序)")
                    .AddOption("old_item_name", ApplicationCommandOptionType.String, "要被替換的原品項【完整名稱】", isRequired: true)
                    .AddOption("new_item_name", ApplicationCommandOptionType.String, "新的品項名稱", isRequired: true)
                    .AddOption("points", ApplicationCommandOptionType.Integer, "新的所需點數 (設為 0 將視為分類標題)", isRequired: true)
                    // 👇 這裡也同步新增類型選擇
                    .AddOption(new SlashCommandOptionBuilder()
                        .WithName("note_type")
                        .WithDescription("選擇備註的標籤類型 (預設為「備註」)")
                        .WithType(ApplicationCommandOptionType.String)
                        .WithRequired(false)
                        .AddChoice("🏷️ 條件", "條件")
                        .AddChoice("📝 備註", "備註"))
                    .AddOption("note", ApplicationCommandOptionType.String, "新的內容 (選填)", isRequired: false)
                    .Build(),
                new SlashCommandBuilder().WithName("edit-menu")
                    .WithDescription("🛍️ 預覽目前的 MENU 草稿，並申請雙人確認發布").Build(),
                new SlashCommandBuilder().WithName("menu").WithDescription("📜 查看目前的兌換點數 MENU (僅自己可見)").Build(),
                new SlashCommandBuilder().WithName("view-admins").WithDescription("👮 查看目前擁有權限的管理員身分組名單").Build(),
                // 在 var commands = new List<ApplicationCommandProperties> 中新增：
                new SlashCommandBuilder().WithName("my-info").WithDescription("👤 查詢自己的點數與近期十筆紀錄").Build(),
                new SlashCommandBuilder().WithName("group-info").WithDescription("👥 查詢已綁定 Roblox 群組中符合資格的總人數").Build(),
                // 👇 替換成下面這段：
                new SlashCommandBuilder()
                    .WithName("service-days")
                    .WithDescription("📅 查詢你或他人在憲兵部門的服役天數")
                    .AddOption("user", ApplicationCommandOptionType.User, "想查詢的對象 (若不填寫則預設查詢自己)", isRequired: false)
                    .Build(),
                // 👇 新增這兩個指令
                new SlashCommandBuilder().WithName("start-event")
                    .WithDescription("🎤 開始紀錄目前所在語音頻道的活動參與時間")
                    .AddOption("reason", ApplicationCommandOptionType.String, "活動名稱或發點原因", isRequired: true)
                    .Build(),
                new SlashCommandBuilder().WithName("end-event")
                    .WithDescription("⏹️ 結束活動並開啟結算面板 (自動過濾待滿5分鐘人員)")
                    .AddOption("points", ApplicationCommandOptionType.Integer, "預定發放點數", isRequired: true)
                    .Build(),
            };
            try { await _client.BulkOverwriteGlobalApplicationCommandsAsync(commands.ToArray()); } catch (Exception ex) { Console.WriteLine(ex.Message); }
        }

        private async Task HandleSlashCommandAsync(SocketSlashCommand command)
        {
            if (!(command.Channel is SocketGuildChannel guildChannel)) return;
            ulong gid = guildChannel.Guild.Id;

            try
            {
                // 將原本的 bool isEphemeral = ... 替換為以下這行 (在最後面加上了 change-menu-item)：
                bool isEphemeral = command.Data.Name == "setup-roca" || command.Data.Name == "unbind-roca" || command.Data.Name == "sync-members" || command.Data.Name == "points" || command.Data.Name == "history" || command.Data.Name == "my-code" || command.Data.Name == "log-channel" || command.Data.Name == "menu" || command.Data.Name == "view-admins" || command.Data.Name == "add-menu-item" || command.Data.Name == "remove-menu-item" || command.Data.Name == "edit-menu" || command.Data.Name == "bind-sheet" || command.Data.Name == "my-info" || command.Data.Name == "group-info" || command.Data.Name == "change-menu-item" || command.Data.Name == "service-days" || command.Data.Name == "start-event" || command.Data.Name == "end-event";
                // 2. 全部統一 Defer (把原本的 if 判斷直接刪除，改成這行)
                await command.DeferAsync(ephemeral: isEphemeral);

                using var db = new BotDbContext(_configuration);
                if (!await db.Database.CanConnectAsync())
                {
                    await command.FollowupAsync("❌ 無法連線到資料庫。");
                    return;
                }

                var botConfig = await db.Configs.FindAsync(gid);

                switch (command.Data.Name)
                {
                    case "setup-roca":
                        {
                            if (!((SocketGuildUser)command.User).GuildPermissions.Administrator) { await command.FollowupAsync("❌ 限管理員執行。"); return; }
                            string rId = (string)command.Data.Options.First(x => x.Name == "roblox_group_id").Value;

                            // 👇 讀取最多七個身分組
                            var role1 = (SocketRole)command.Data.Options.FirstOrDefault(x => x.Name == "admin_role_1")?.Value;
                            var role2 = (SocketRole)command.Data.Options.FirstOrDefault(x => x.Name == "admin_role_2")?.Value;
                            var role3 = (SocketRole)command.Data.Options.FirstOrDefault(x => x.Name == "admin_role_3")?.Value;
                            var role4 = (SocketRole)command.Data.Options.FirstOrDefault(x => x.Name == "admin_role_4")?.Value;
                            var role5 = (SocketRole)command.Data.Options.FirstOrDefault(x => x.Name == "admin_role_5")?.Value;
                            var role6 = (SocketRole)command.Data.Options.FirstOrDefault(x => x.Name == "admin_role_6")?.Value;
                            var role7 = (SocketRole)command.Data.Options.FirstOrDefault(x => x.Name == "admin_role_7")?.Value;

                            var roleIds = new List<ulong>();
                            if (role1 != null) roleIds.Add(role1.Id);
                            if (role2 != null) roleIds.Add(role2.Id);
                            if (role3 != null) roleIds.Add(role3.Id);
                            if (role4 != null) roleIds.Add(role4.Id);
                            if (role5 != null) roleIds.Add(role5.Id);
                            if (role6 != null) roleIds.Add(role6.Id);
                            if (role7 != null) roleIds.Add(role7.Id);
                            string roleIdsStr = string.Join(",", roleIds); // 轉成逗號分隔字串

                            string newCode = Guid.NewGuid().ToString("N").Substring(0, 5).ToUpper();

                            if (botConfig == null)
                            {
                                botConfig = new BotConfig { GuildId = gid, RobloxGroupId = rId, AdminRoleId = role1.Id, AdminRoleIds = roleIdsStr, ServerCode = newCode };
                                db.Configs.Add(botConfig);
                            }
                            else
                            {
                                botConfig.RobloxGroupId = rId; botConfig.AdminRoleId = role1.Id; botConfig.AdminRoleIds = roleIdsStr;
                                if (string.IsNullOrEmpty(botConfig.ServerCode)) botConfig.ServerCode = newCode;
                            }
                            await db.SaveChangesAsync();

                            await command.FollowupAsync($"✅ 已綁定 Roblox 群組 `{rId}`。\n🔑 **本伺服器專屬資料庫編號：`{botConfig.ServerCode}`** (請妥善保存，交給國防部設定 Admin 頻道使用)。\n正在背景抓取群組成員，請稍候...");
                            var syncResult = await SyncGroupMembersAsync(db, gid, rId);
                            await command.Channel.SendMessageAsync($"🔄 群組名單同步完成！\n> 🟢 新增了 **{syncResult.added}** 位新成員 (預設為 0 點)\n> 🔴 移除了 **{syncResult.removed}** 位已退群或不符資格成員。");
                            break;
                        }
                    case "start-event":
                        {
                            if (botConfig == null) { await command.FollowupAsync("❌ 請先設定機器人。"); return; }
                            var exec = (SocketGuildUser)command.User;
                            if (!IsAdmin(exec, botConfig) && exec.Id != guildChannel.Guild.OwnerId) { await command.FollowupAsync("❌ 權限不足。"); return; }

                            if (exec.VoiceChannel == null)
                            {
                                await command.FollowupAsync("❌ 您必須先進入一個「語音頻道」，才能對該頻道開始紀錄！");
                                return;
                            }

                            ulong vcId = exec.VoiceChannel.Id;
                            if (_activeEvents.ContainsKey(vcId))
                            {
                                await command.FollowupAsync($"❌ 頻道 <#{vcId}> 已經在紀錄中了！請先使用 `/end-event` 結束它。");
                                return;
                            }

                            string reason = (string)command.Data.Options.First(x => x.Name == "reason").Value;

                            var newEvt = new VoiceEventSession { GuildId = gid, ChannelId = vcId, AdminId = exec.Id, Reason = reason, StartTime = Program.GetTaipeiTime() };

                            // 把當下已經在頻道內的人先加進去，當作 0 秒起算
                            foreach (var u in exec.VoiceChannel.Users)
                            {
                                newEvt.Participants[u.Id] = new VoiceParticipant { LastJoinTime = Program.GetTaipeiTime() };
                                newEvt.ActionLogs.Add($"[{Program.GetTaipeiTime():HH:mm:ss}] 🟢 {u.Username} (原先已在頻道內)");
                            }

                            _activeEvents[vcId] = newEvt;
                            await command.FollowupAsync($"🎤 **活動開始紀錄！**\n> 📍 **頻道：** <#{vcId}>\n> 📝 **原因：** {reason}\n> ⏱️ **時間：** {newEvt.StartTime:HH:mm:ss}\n> *(機器人將自動統計進出時間，請於活動結束時在此輸入 `/end-event`)*");
                            break;
                        }

                    case "end-event":
                        {
                            if (botConfig == null) { await command.FollowupAsync("❌ 請先設定機器人。"); return; }
                            var exec = (SocketGuildUser)command.User;
                            if (!IsAdmin(exec, botConfig) && exec.Id != guildChannel.Guild.OwnerId) { await command.FollowupAsync("❌ 權限不足。"); return; }

                            if (exec.VoiceChannel == null) { await command.FollowupAsync("❌ 您必須待在要結束紀錄的「語音頻道」內！"); return; }

                            ulong vcId = exec.VoiceChannel.Id;
                            if (!_activeEvents.TryGetValue(vcId, out var evt))
                            {
                                await command.FollowupAsync($"❌ 頻道 <#{vcId}> 目前沒有正在紀錄的活動。");
                                return;
                            }

                            int points = Convert.ToInt32((long)command.Data.Options.First(x => x.Name == "points").Value);
                            var now = Program.GetTaipeiTime();

                            // 結算所有人的最終停留時間
                            foreach (var kvp in evt.Participants)
                            {
                                if (kvp.Value.LastJoinTime.HasValue)
                                {
                                    kvp.Value.TotalSeconds += (now - kvp.Value.LastJoinTime.Value).TotalSeconds;
                                    kvp.Value.LastJoinTime = null;
                                }
                            }

                            // 🏆 篩選出大於等於 5 分鐘 (300 秒) 的人
                            var qualifiedIds = evt.Participants.Where(x => x.Value.TotalSeconds >= 300).Select(x => x.Key).ToHashSet();

                            string sessId = Guid.NewGuid().ToString("N").Substring(0, 8);
                            _pendingDistributions[sessId] = new PendingEventDistribution
                            {
                                GuildId = gid,
                                Points = points,
                                Reason = evt.Reason,
                                AdminId = exec.Id,
                                FinalUserIds = qualifiedIds,
                                ActionLogs = evt.ActionLogs,
                                StartTime = evt.StartTime, // 👈 帶入開始時間
                                EndTime = now              // 👈 帶入剛剛取得的結束時間
                            };

                            _activeEvents.Remove(vcId); // 停止追蹤此頻道
                            await UpdateEventDashboardAsync(command, sessId, guildChannel.Guild, isFirst: true);
                            break;
                        }
                    case "sync-members":
                        {
                            if (botConfig == null) { await command.FollowupAsync("❌ 請先使用 `/setup-roca` 綁定群組。"); return; }
                            if (!IsAdmin((SocketGuildUser)command.User, botConfig)) { await command.FollowupAsync("❌ 權限不足。"); return; }

                            await command.FollowupAsync("⏳ 正在與 Roblox 同步群組成員名單...");
                            var syncResult = await SyncGroupMembersAsync(db, gid, botConfig.RobloxGroupId);
                            // 在背景非同步執行試算表更新，不影響 Discord 機器人的回應速度
                            _ = UpdateGoogleSheetAsync(gid);
                            await command.FollowupAsync($"✅ 同步完成！\n> 🟢 本次共新增了 **{syncResult.added}** 位新成員\n> 🔴 移除了 **{syncResult.removed}** 位已退群或不符資格成員。");
                            break;
                        }

                    case "my-code":
                        {
                            if (botConfig == null || string.IsNullOrEmpty(botConfig.ServerCode))
                            {
                                await command.FollowupAsync("❌ 本伺服器尚未完成設定。請先使用 `/setup-roca` 進行綁定。");
                                break;
                            }

                            var myCodeExec = (SocketGuildUser)command.User;
                            bool isCodeAdmin = IsAdmin(myCodeExec, botConfig);

                            if (!isCodeAdmin)
                            {
                                await command.FollowupAsync("❌ 權限不足！只有管理員可以查詢伺服器專屬編號。");
                                return;
                            }

                            await command.FollowupAsync($"🔑 **本伺服器專屬資料庫編號為：** `{botConfig.ServerCode}`\n> 請將此編號交給國防部，以便國防部進行跨伺服器監控設定。");
                            break;
                        }
                    case "log-channel":
                        {
                            if (botConfig == null) { await command.FollowupAsync("❌ 請先使用 `/setup-roca` 完成基本設定。"); break; }

                            var execUser = (SocketGuildUser)command.User;
                            // 限制權限：只有伺服器管理員，或是擁有綁定的 Admin 身分組的人可以設定
                            if (!IsAdmin((SocketGuildUser)command.User, botConfig)) 
                            { 
                                await command.FollowupAsync("❌ 權限不足。"); 
                                return; 
                            }

                            var channelOpt = (SocketChannel)command.Data.Options.First().Value;
                            if (!(channelOpt is ITextChannel))
                            {
                                await command.FollowupAsync("❌ 格式錯誤，請選擇一個「文字頻道」！");
                                return;
                            }

                            // 寫入資料庫
                            botConfig.MasterLogChannelId = channelOpt.Id;
                            await db.SaveChangesAsync();
                            await command.FollowupAsync($"✅ **已成功將本伺服器的點數變動推播頻道設為：** <#{channelOpt.Id}>\n> 往後只要有點數發放、扣除或撤銷，都會自動傳送到該頻道！");
                            break;
                        }
                    case "viewall":
                        {
                            var all = await db.UserPoints.Where(u => u.GuildId == gid).OrderByDescending(u => u.Points).ToListAsync();
                            if (!all.Any()) { await command.FollowupAsync("📭 尚無資料。"); break; }

                            var chunks = new List<string>();
                            var currentChunk = new StringBuilder("##  點數總覽 (所有成員)\n```ansi\n");

                            int rankIndex = 1;
                            foreach (var u in all)
                            {
                                string rankRaw = $"[{rankIndex}]".PadRight(5);
                                string nameRaw = u.RobloxUsername.PadRight(20);
                                string pointRaw = $"{u.Points} 點".PadLeft(8);

                                string rankStr = $"\u001b[34m{rankRaw}\u001b[0m";
                                string pointStr = $"\u001b[32m{pointRaw}\u001b[0m";

                                string line = $"{rankStr} {nameRaw} ➔ {pointStr}\n";

                                if (currentChunk.Length + line.Length > 1900)
                                {
                                    currentChunk.AppendLine("```");
                                    chunks.Add(currentChunk.ToString());
                                    currentChunk.Clear();
                                    currentChunk.AppendLine("```ansi\n");
                                }
                                currentChunk.Append(line);
                                rankIndex++;
                            }
                            if (currentChunk.Length > 0)
                            {
                                currentChunk.AppendLine("```");
                                chunks.Add(currentChunk.ToString());
                            }
                            // 👇 加這行：在最後一頁資料的結尾加上時間
                            chunks[chunks.Count - 1] += $"\n>  *數據抓取時間：{Program.GetTaipeiTime():yyyy-MM-dd HH:mm:ss}*";

                            bool isFirst = true;
                            foreach (var chunk in chunks)
                            {
                                if (isFirst) { await command.FollowupAsync(chunk); isFirst = false; }
                                else { await command.Channel.SendMessageAsync(chunk); }
                            }
                            break;
                        }
                    case "clear-all-data":
                        {
                            if (botConfig == null) { await command.FollowupAsync("❌ 請先設定機器人。"); return; }
                            if (!IsAdmin((SocketGuildUser)command.User, botConfig)) { await command.FollowupAsync("❌ 限管理員執行。"); return; }

                            int adminCount = GetAdminCount(guildChannel.Guild, botConfig);
                            bool requiresDual = adminCount >= 2;

                            string btnId = requiresDual ? $"clear_step1_{gid}" : $"clear_single_{gid}";
                            string btnText = requiresDual ? "⚠️ 第一位管理員確認" : "⚠️ 確認清空 (單人管理員)";

                            var btns = new ComponentBuilder()
                                .WithButton(btnText, btnId, ButtonStyle.Danger)
                                .WithButton("取消", $"clear_cancel_{gid}", ButtonStyle.Secondary);

                            string msg = requiresDual ? "\n> 執行此動作需要 **兩名管理員** 共同確認，大家都會看到此操作。" : $"\n> ⚠️ 系統偵測到目前管理員僅有 {adminCount} 人，已切換為**單人快速確認**。";
                            await command.FollowupAsync($"⚠️ **【資料清空請求】** 確定要清空所有資料嗎？{msg}", components: btns.Build());
                            break;
                        }

                    case "unbind-roca":
                        {
                            if (botConfig == null) { await command.FollowupAsync("⚠️ 本伺服器尚未進行綁定。"); return; }
                            if (!IsAdmin((SocketGuildUser)command.User, botConfig)) { await command.FollowupAsync("❌ 限管理員執行。"); return; }

                            int adminCount = GetAdminCount(guildChannel.Guild, botConfig);
                            bool requiresDual = adminCount >= 2;

                            string btnId = requiresDual ? $"unbind_step1_{gid}" : $"unbind_single_{gid}";
                            string btnText = requiresDual ? "⚠️ 第一位管理員確認" : "⚠️ 確認解除 (單人管理員)";

                            var btns = new ComponentBuilder()
                                .WithButton(btnText, btnId, ButtonStyle.Danger)
                                .WithButton("取消", $"unbind_cancel_{gid}", ButtonStyle.Secondary);

                            string msg = requiresDual ? "\n> 執行此動作需要 **兩名管理員** 共同確認。" : $"\n> ⚠️ 系統偵測到目前管理員僅有 {adminCount} 人，已切換為**單人快速確認**。";
                            await command.FollowupAsync($"⚠️ **【解除綁定請求】** 確定要解除綁定嗎？這將會移除伺服器的單位對應設定！{msg}", components: btns.Build());
                            break;
                        }

                    case "edit-admin":
                        {
                            if (botConfig == null) { await command.FollowupAsync("⚠️ 請先使用 `/setup-roca` 設定機器人。"); return; }
                            if (!IsAdmin((SocketGuildUser)command.User, botConfig)) { await command.FollowupAsync("❌ 限管理員執行。"); return; }

                            string action = (string)command.Data.Options.First(x => x.Name == "action").Value;
                            var role = (SocketRole)command.Data.Options.First(x => x.Name == "role").Value;
                            var currentRoles = botConfig.AdminRoleIds?.Split(',').Where(s => !string.IsNullOrEmpty(s)).Select(ulong.Parse).ToList() ?? new List<ulong>();

                            if (action == "remove" && currentRoles.Count <= 1 && currentRoles.Contains(role.Id)) { await command.FollowupAsync("❌ 操作無效！您必須至少保留一個管理員身分組。"); return; }
                            if (action == "add" && currentRoles.Contains(role.Id)) { await command.FollowupAsync("⚠️ 該身分組已經在管理員名單中了。"); return; }

                            string actionText = action == "add" ? "新增" : "移除";
                            int adminCount = GetAdminCount(guildChannel.Guild, botConfig);
                            bool requiresDual = adminCount >= 2;

                            string btnId = requiresDual ? $"editadm_s1_{gid}_{action}_{role.Id}" : $"editadm_single_{gid}_{action}_{role.Id}";
                            string btnText = requiresDual ? "⚠️ 第一位管理員確認" : "⚠️ 確認變更 (單人管理員)";

                            var btns = new ComponentBuilder()
                                .WithButton(btnText, btnId, ButtonStyle.Danger)
                                .WithButton("取消", $"editadm_cancel_{gid}", ButtonStyle.Secondary);

                            string msg = requiresDual ? "\n> 執行此動作需要 **兩名管理員** 共同確認。" : $"\n> ⚠️ 系統偵測到目前管理員僅有 {adminCount} 人，已切換為**單人快速確認**。";
                            await command.FollowupAsync($"⚠️ **【權限變更請求】** 確定要 **{actionText}** <@&{role.Id}> 的管理員權限嗎？{msg}", components: btns.Build());
                            break;
                        }

                    case "edit-menu": // 👈 這裡永久拔除雙管理員確認
                        {
                            if (botConfig == null) { await command.FollowupAsync("❌ 請先設定機器人。"); return; }
                            if (!IsAdmin((SocketGuildUser)command.User, botConfig)) { await command.FollowupAsync("❌ 限管理員執行。"); return; }

                            var items = await db.RewardMenuItems.Where(x => x.GuildId == gid).OrderBy(x => x.Id).ToListAsync();
                            string preview = FormatMenuItems(items);

                            var btns = new ComponentBuilder()
                                .WithButton("✅ 草稿無誤，立即發布", $"menuok_single_{gid}", ButtonStyle.Success)
                                .WithButton("取消 (繼續修改)", $"menuno_{gid}", ButtonStyle.Secondary);

                            await command.FollowupAsync($"👁️ **【草稿預覽】** 以下是目前草稿的排版結果：\n{preview}\n> 若確認無誤，請點擊下方按鈕直接發布更新。", components: btns.Build());
                            break;
                        }

                    case "group-info":
                        {
                            if (botConfig == null || string.IsNullOrEmpty(botConfig.RobloxGroupId)) { await command.FollowupAsync("❌ 請先設定機器人。"); return; }
                            await command.FollowupAsync("⏳ 正在向 Roblox 查詢資料，請稍候...");

                            int maxRank = 255;
                            string unitName = "未知單位"; // 👈 新增變數用來儲存單位名稱

                            switch (botConfig.RobloxGroupId)
                            {
                                case "13549943": maxRank = 239; unitName = "國防部憲兵指揮部"; break; // 👈 順便記錄單位名稱
                                case "13662982": maxRank = 120; unitName = "陸軍裝甲第五八四旅"; break;
                                case "16223475": maxRank = 55; unitName = "陸軍航空特戰指揮部"; break;
                                default: unitName = "自訂單位"; break;
                            }
                            int minRank = 1;
                            int count = 0;
                            string cursor = "";
                            bool hasMore = true;

                            try
                            {
                                while (hasMore)
                                {
                                    string url = $"https://groups.roblox.com/v1/groups/{botConfig.RobloxGroupId}/users?limit=100";
                                    if (!string.IsNullOrEmpty(cursor)) url += $"&cursor={cursor}";
                                    var res = await _http.GetAsync(url);
                                    if (!res.IsSuccessStatusCode) break;

                                    var json = await res.Content.ReadAsStringAsync();
                                    using var doc = JsonDocument.Parse(json);
                                    var root = doc.RootElement;
                                    var data = root.GetProperty("data");
                                    foreach (var item in data.EnumerateArray())
                                    {
                                        int rank = item.GetProperty("role").GetProperty("rank").GetInt32();
                                        if (rank >= minRank && rank <= maxRank) count++;
                                    }
                                    if (root.TryGetProperty("nextPageCursor", out var nextCursor) && nextCursor.ValueKind == JsonValueKind.String)
                                    {
                                        cursor = nextCursor.GetString();
                                        await Task.Delay(500); // 防頻頸
                                    }
                                    else hasMore = false;
                                }
                                // 👇 修改這行：隱藏 Rank 條件，只顯示部門名稱與總人數
                                await command.FollowupAsync($"👥 **部門人數查詢**\n> 您所在的部門 **[{unitName}]** 目前總人數為：**{count}** 人。\n> *(此數據為即時向 Roblox 伺服器查詢之結果)*");
                            }
                            catch (Exception ex)
                            {
                                await command.FollowupAsync($"❌ 查詢時發生錯誤，請稍後再試: {ex.Message}");
                            }
                            break;
                        }

                    case "my-info":
                        {
                            var targetUser = (SocketGuildUser)command.User;
                            string targetName = targetUser.Nickname ?? targetUser.Username;
                            // 自動識別 `]` 後面的文字當作 Roblox Name
                            if (targetName.Contains("]")) targetName = targetName.Substring(targetName.LastIndexOf(']') + 1).Trim();

                            var userPoint = await db.UserPoints.FirstOrDefaultAsync(u => u.GuildId == gid && u.RobloxUsername.ToLower() == targetName.ToLower());
                            int currentPoints = userPoint != null ? userPoint.Points : 0;

                            var logs = await db.PointLogs.Where(l => l.GuildId == gid && l.RobloxUsername.ToLower() == targetName.ToLower() && !l.IsDeleted).OrderByDescending(l => l.Timestamp).Take(10).ToListAsync();

                            var sb = new StringBuilder($"### 👤 **{targetName}** 的個人資訊\n");
                            sb.AppendLine($"> **目前可用點數：** `{currentPoints}` 點\n");

                            if (!logs.Any())
                            {
                                sb.AppendLine("📭 您尚無近期的點數變動紀錄。");
                            }
                            else
                            {
                                sb.AppendLine("📜 **近期紀錄 (最多顯示十筆)：**\n```ansi\n");
                                foreach (var l in logs)
                                {
                                    string idRaw = $"[ID: {l.Id}]".PadRight(9);
                                    string timeStr = l.Timestamp.ToString("MM/dd HH:mm");
                                    string ptsRaw = (l.PointsAdded > 0 ? $"+{l.PointsAdded}" : l.PointsAdded.ToString()).PadLeft(5);
                                    string adminStr = l.AdminName.PadRight(12);
                                    string idStr = $"\u001b[34m{idRaw}\u001b[0m";
                                    string ptsStr = (l.PointsAdded > 0 ? "\u001b[32m" : "\u001b[31m") + $"{ptsRaw}\u001b[0m";
                                    sb.AppendLine($"{idStr} \u001b[30m{timeStr}\u001b[0m | ➔ {ptsStr} 點 | 登記: {adminStr} | 原因: {l.Reason}");
                                }
                                sb.AppendLine("```");
                            }
                            sb.AppendLine($"\n> *數據抓取時間：{Program.GetTaipeiTime():yyyy-MM-dd HH:mm:ss}*");
                            await command.FollowupAsync(sb.ToString());
                            break;
                        }
                    case "addpoint":
                        {
                            if (botConfig == null) { await command.FollowupAsync("❌ 未設定。請先使用 /setup-roca"); break; }
                            var exec = (SocketGuildUser)command.User;
                            if (!IsAdmin(exec, botConfig) && exec.Id != guildChannel.Guild.OwnerId) { await command.FollowupAsync("❌ 權限不足。"); return; }

                            var user = (SocketGuildUser)command.Data.Options.First(x => x.Name == "user").Value;
                            int pts = Convert.ToInt32((long)command.Data.Options.First(x => x.Name == "points").Value);
                            string reason = (string)command.Data.Options.First(x => x.Name == "reason").Value;
                            string name = user.Nickname ?? user.Username;
                            if (name.Contains("]")) name = name.Substring(name.LastIndexOf(']') + 1).Trim();

                            if (!await VerifyUserInRobloxGroup(name, botConfig.RobloxGroupId)) { await command.FollowupAsync($"❌ 玩家 `{name}` 不在指定的 Roblox 群組內，或名稱不相符。"); break; }

                            var rec = await db.UserPoints.FirstOrDefaultAsync(u => u.GuildId == gid && u.RobloxUsername.ToLower() == name.ToLower());
                            if (rec == null) { rec = new UserPoint { GuildId = gid, RobloxUsername = name, Points = 0 }; db.UserPoints.Add(rec); }

                            rec.Points += pts;

                            var newLog = new PointLog { GuildId = gid, RobloxUsername = name, AdminName = command.User.Username, PointsAdded = pts, Reason = reason, Timestamp = Program.GetTaipeiTime() };
                            db.PointLogs.Add(newLog);
                            await db.SaveChangesAsync();
                            // 在背景非同步執行試算表更新，不影響 Discord 機器人的回應速度
                            _ = UpdateGoogleSheetAsync(gid);

                            // 【修改後的訊息格式：移除反引號 ` ，改用粗體 ** 】
                            string addMsg = $"> **[➕點數發放]** 登記人：**{command.User.Username}**\n" +
                                $"> **✅發放 {pts} 點，給 {name}**\n" +
                                $"> **目前總計**：**{rec.Points}** 點\n" +
                                $"> **備註：{reason}**\n" +
                                $"> **紀錄編號**：{newLog.Id}";
                            await command.FollowupAsync(addMsg);

                            bool isAnomaly = pts >= 100; // 判斷是否大於等於 100
                            _ = BroadcastToAdminChannelsAsync(gid, $"負責人 **{command.User.Username}** 已發放點數\n" +
                                        $">  **被登記人：{name}獲得 {pts} 點**\n" +
                                        $">  **目前總計：{rec.Points} 點**\n" +
                                        $">  **備註：{reason}**\n" +
                                        $">  **紀錄編號**：**{newLog.Id}** (若需撤銷請使用 /del-record)", isAnomaly);
                            break;
                        }
                    case "removepoint":
                        {
                            if (botConfig == null) { await command.FollowupAsync("❌ 未設定。請先使用 /setup-roca"); break; }
                            var exec = (SocketGuildUser)command.User;
                            if (!IsAdmin(exec, botConfig) && exec.Id != guildChannel.Guild.OwnerId) { await command.FollowupAsync("❌ 權限不足。"); return; }

                            var user = (SocketGuildUser)command.Data.Options.First(x => x.Name == "user").Value;
                            int pts = Convert.ToInt32((long)command.Data.Options.First(x => x.Name == "points").Value);
                            string reason = (string)command.Data.Options.First(x => x.Name == "reason").Value;

                            if (pts <= 0) { await command.FollowupAsync("❌ 扣除的點數必須大於 0。"); break; }

                            string name = user.Nickname ?? user.Username;
                            if (name.Contains("]")) name = name.Substring(name.LastIndexOf(']') + 1).Trim();

                            if (!await VerifyUserInRobloxGroup(name, botConfig.RobloxGroupId)) { await command.FollowupAsync($"❌ 玩家 `{name}` 不在指定的 Roblox 群組內，或名稱不相符。"); break; }

                            var rec = await db.UserPoints.FirstOrDefaultAsync(u => u.GuildId == gid && u.RobloxUsername.ToLower() == name.ToLower());

                            // 檢查玩家有沒有足夠的點數可以扣
                            if (rec == null || rec.Points < pts)
                            {
                                int currentPts = rec == null ? 0 : rec.Points;
                                await command.FollowupAsync($"❌ **{name}** 的點數不足！目前僅有 **{currentPts}** 點，無法扣除 **{pts}** 點。");
                                break;
                            }

                            // 執行扣點
                            rec.Points -= pts;

                            // 紀錄寫入資料庫，注意這裡的 PointsAdded 存為 -pts (負數)
                            var newLog = new PointLog { GuildId = gid, RobloxUsername = name, AdminName = command.User.Username, PointsAdded = -pts, Reason = reason, Timestamp = Program.GetTaipeiTime() };
                            db.PointLogs.Add(newLog);
                            await db.SaveChangesAsync();
                            // 在背景非同步執行試算表更新，不影響 Discord 機器人的回應速度
                            _ = UpdateGoogleSheetAsync(gid);

                            string removeMsg = $">  **[➖點數扣除/兌換]** 登記人：**{command.User.Username}**\n" +
                                               $">  **✅成功扣除 {pts} 點，自 {name}**\n" +
                                               $">  **目前剩餘**：**{rec.Points}** 點\n" +
                                               $">  **備註：{reason}** \n" +
                                               $">  **紀錄編號**：{newLog.Id}";
                            await command.FollowupAsync(removeMsg);

                            // 推播給總部
                            bool isAnomaly = pts >= 100; // 判斷是否大於等於 100
                            _ = BroadcastToAdminChannelsAsync(gid, $"負責人 **{command.User.Username}** 已扣除點數\n" +
                                        $">  被登記人：**{name}** 扣除 **{pts}** 點\n" +
                                        $">  目前剩餘：**{rec.Points}** 點\n" +
                                        $">  備註：{reason}\n" +
                                        $">  紀錄編號：**{newLog.Id}**", isAnomaly);
                            break;
                        }
                    case "service-days":
                        {
                            // 1. 檢查該伺服器是否綁定了憲兵群組 (ID: 13549943)
                            if (botConfig == null || botConfig.RobloxGroupId != "13549943")
                            {
                                await command.FollowupAsync("❌ **權限拒絕：** 此為 **憲兵部門** 的專屬功能！本伺服器尚未綁定憲兵群組，無法使用。");
                                return;
                            }

                            // 👇 2. 判斷是否有選擇特定使用者，若沒有則預設為「執行指令的人」
                            var targetUserOption = command.Data.Options.FirstOrDefault(x => x.Name == "user");
                            SocketGuildUser targetUser = targetUserOption != null ? (SocketGuildUser)targetUserOption.Value : (SocketGuildUser)command.User;

                            string targetName = targetUser.Nickname ?? targetUser.Username;

                            // 3. 自動辨識並擷取後方的 Roblox Name
                            if (targetName.Contains("]"))
                                targetName = targetName.Substring(targetName.LastIndexOf(']') + 1).Trim();

                            // 4. 驗證該名單是否在資料庫
                            var userPoint = await db.UserPoints.FirstOrDefaultAsync(u => u.GuildId == gid && u.RobloxUsername.ToLower() == targetName.ToLower());
                            if (userPoint == null)
                            {
                                string errMsg = targetUserOption != null ? $"❌ 找不到玩家 `{targetName}` 的資料，請確認該員是否在綁定的 Roblox 群組名單內。" : $"❌ 找不到您的資料，請確認您是否在綁定的 Roblox 群組名單內。";
                                await command.FollowupAsync(errMsg);
                                return;
                            }

                            // 5. 取得指定的「🧾課程結果-class-result」頻道
                            ulong classResultChannelId = 1325441624704286760;
                            var channel = _client.GetChannel(classResultChannelId) as ITextChannel;
                            if (channel == null)
                            {
                                await command.FollowupAsync("❌ 發生錯誤：找不到指定的課程結果頻道，請確認機器人是否已被邀請至該頻道並具備「讀取訊息歷史」權限。");
                                return;
                            }

                            await command.FollowupAsync($"⏳ 正在翻閱 `{targetName}` 的結訓紀錄檔案，這可能需要幾秒鐘的時間...");

                            IMessage foundMessage = null;

                            // 6. 往前翻找歷史訊息 (設定最多往回找 2000 筆以保護效能)
                            await foreach (var batch in channel.GetMessagesAsync(2000))
                            {
                                foreach (var msg in batch)
                                {
                                    // 檢查訊息內是否包含「通過」以及該使用者的 ID 或 Roblox Name
                                    if (msg.Content.Contains("通過") && (msg.Content.Contains(targetUser.Id.ToString()) || msg.Content.Contains(targetName)))
                                    {
                                        // 為了避免在「不通過」被 Ping 到，逐行嚴格檢查
                                        var lines = msg.Content.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                                        bool isPassed = false;
                                        foreach (var line in lines)
                                        {
                                            if (line.Contains("通過") && (line.Contains(targetUser.Id.ToString()) || line.Contains(targetName)))
                                            {
                                                isPassed = true;
                                                break;
                                            }
                                        }

                                        if (isPassed)
                                        {
                                            foundMessage = msg;
                                            break;
                                        }
                                    }
                                }
                                if (foundMessage != null) break; // 找到最新的一筆就停止搜尋
                            }

                            if (foundMessage == null)
                            {
                                // 👇 修改這裡：把 command.Channel.SendMessageAsync 改成 command.FollowupAsync
                                await command.FollowupAsync($"📭 {targetUser.Mention} 在 <#{classResultChannelId}> 頻道中，找不到結訓通過紀錄。\n> *(提示：機器人目前最多往前回溯 2000 筆紀錄，若是較久以前的紀錄可能需要重新補登)*");
                                return;
                            }
                            // 7. 換算台北時間並計算天數
                            DateTime msgTimeUtc = foundMessage.Timestamp.UtcDateTime;
                            DateTime msgTimeTaipei = TimeZoneInfo.ConvertTimeFromUtc(msgTimeUtc, TimeZoneInfo.FindSystemTimeZoneById("Taipei Standard Time"));
                            DateTime nowTaipei = Program.GetTaipeiTime();

                            // 用 Date 屬性相減，確保是計算「日曆天數」的差距
                            int days = (int)(nowTaipei.Date - msgTimeTaipei.Date).TotalDays;

                            // 👇 取得打指令者 (執行者) 的名稱，並過濾掉 Roblox 綴詞
                            var executorUser = (SocketGuildUser)command.User;
                            string executorName = executorUser.Nickname ?? executorUser.Username;
                            if (executorName.Contains("]"))
                                executorName = executorName.Substring(executorName.LastIndexOf(']') + 1).Trim();

                            // 👇 判斷是否為查詢自己，動態切換稱呼與標題
                            bool isSelf = targetUser.Id == command.User.Id;
                            string pronoun = isSelf ? "您" : "該員";

                            string titleMsg = isSelf
                                ? $" **服役天數查詢：{targetName}**"
                                : $" **{executorName}** 替 **{targetName}** 查詢的服役天數：";

                            // 👇 發送最終的隱藏訊息 (FollowupAsync)
                            await command.FollowupAsync($"{titleMsg}\n> 自 **{msgTimeTaipei:yyyy年MM月dd日}** 結訓通過起算\n> {pronoun}在部門已經服役了 **{days}** 天！\n");
                            break;
                        }
                    case "status":
                        {
                            int latency = _client.Latency;
                            bool isDbConnected = await db.Database.CanConnectAsync();
                            string dbStatusMsg = isDbConnected ? "✅ 正常連線" : "❌ 無法連線";

                            string statusMessage = $"🟢 **ROCA Point Bot 運作正常！**\n" +
                                                   $">  **Discord 延遲 (Ping):** `{latency} ms`\n" +
                                                   $">  **資料庫狀態:** {dbStatusMsg}\n" +
                                                   $">  **主機目前時間:** `{Program.GetTaipeiTime():yyyy-MM-dd HH:mm:ss}`";

                            await command.FollowupAsync(statusMessage);
                            break;
                        }

                    case "history":
                        {
                            var hUser = (SocketGuildUser)command.Data.Options.First().Value;
                            string hName = hUser.Nickname ?? hUser.Username;
                            if (hName.Contains("]")) hName = hName.Substring(hName.LastIndexOf(']') + 1).Trim();
                            var logs = await db.PointLogs.Where(l => l.GuildId == gid && l.RobloxUsername.ToLower() == hName.ToLower() && !l.IsDeleted).OrderByDescending(l => l.Timestamp).Take(10).ToListAsync();
                            if (!logs.Any()) { await command.FollowupAsync("📭 查無紀錄。"); break; }

                            var sb = new StringBuilder($"###  **{hName}** 的近期紀錄\n```ansi\n");
                            foreach (var l in logs)
                            {
                                string idRaw = $"[ID: {l.Id}]".PadRight(9);
                                string timeStr = l.Timestamp.ToString("MM/dd HH:mm");
                                // 自動判斷正負號
                                string ptsRaw = (l.PointsAdded > 0 ? $"+{l.PointsAdded}" : l.PointsAdded.ToString()).PadLeft(5);
                                string adminStr = l.AdminName.PadRight(12);

                                string idStr = $"\u001b[34m{idRaw}\u001b[0m";
                                // 正數維持綠色(32m)，負數改為紅色(31m)
                                string ptsStr = (l.PointsAdded > 0 ? "\u001b[32m" : "\u001b[31m") + $"{ptsRaw}\u001b[0m";

                                sb.AppendLine($"{idStr} \u001b[30m{timeStr}\u001b[0m | ➔ {ptsStr} 點 | 登記: {adminStr} | 原因: {l.Reason}");
                            }
                            sb.AppendLine("```");
                            sb.AppendLine($"\n> *數據抓取時間：{Program.GetTaipeiTime():yyyy-MM-dd HH:mm:ss}*");
                            await command.FollowupAsync(sb.ToString());
                            break;
                        }

                    case "del-record":
                        {
                            if (botConfig == null) { await command.FollowupAsync("⚠️ 請先設定。"); return; }
                            if (!IsAdmin((SocketGuildUser)command.User, botConfig)) { await command.FollowupAsync("❌ 權限不足。"); return; }

                            // 接收 ID 與必填的原因
                            int logId = Convert.ToInt32((long)command.Data.Options.First(x => x.Name == "id").Value);
                            string delReason = (string)command.Data.Options.First(x => x.Name == "reason").Value;

                            var targetLog = await db.PointLogs.FirstOrDefaultAsync(l => l.Id == logId && l.GuildId == gid);

                            if (targetLog == null || targetLog.IsDeleted) { await command.FollowupAsync("❌ 找不到該紀錄或已經被撤銷過了。"); break; }

                            var uPoint = await db.UserPoints.FirstOrDefaultAsync(u => u.GuildId == gid && u.RobloxUsername.ToLower() == targetLog.RobloxUsername.ToLower());

                            int currentPoints = 0;

                            if (uPoint != null)
                            {
                                uPoint.Points = Math.Max(0, uPoint.Points - targetLog.PointsAdded);
                                currentPoints = uPoint.Points;
                            }

                            targetLog.IsDeleted = true;
                            // 將撤銷原因加到資料庫原有的 Reason 後面，方便未來追蹤
                            targetLog.Reason += $" (已由 {command.User.Username} 撤銷，原因：{delReason})";

                            await db.SaveChangesAsync();
                            // 在背景非同步執行試算表更新，不影響 Discord 機器人的回應速度
                            _ = UpdateGoogleSheetAsync(gid);
                            // 顯示在前端頻道給大家看的訊息
                            string delMsg = $" **[紀錄撤銷]** 負責人：**{command.User.Username}**\n" +
                                            $"> 成功撤銷了紀錄 #{logId}\n" +
                                            $">  **被扣點人：{targetLog.RobloxUsername}**\n" +
                                            $">  **已扣除：{targetLog.PointsAdded} 點**\n" +
                                            $">  **扣除原因：{delReason}**\n" +
                                            $">  **目前剩餘：{currentPoints} 點**";

                            await command.FollowupAsync(delMsg);

                            // 🚨 刪除紀錄屬於敏感操作，一律視同異常並觸發 Ping 第一管理員身分組
                            _ = BroadcastToAdminChannelsAsync(gid, $"負責人 **{command.User.Username}** 撤銷了紀錄 `#{logId}`，撤銷了 {targetLog.PointsAdded} 點的變動。\n>  人員：{targetLog.RobloxUsername}\n>  撤銷原因：{delReason}\n>  剩餘：{currentPoints}點", true);
                            break;
                        }

                    case "daily-records":
                        {
                            // 👇 新增這兩行：檢查伺服器是否已設定，並驗證執行者是否具備管理員權限
                            if (botConfig == null) { await command.FollowupAsync("❌ 請先使用 `/setup-roca` 完成基本設定。"); break; }
                            if (!IsAdmin((SocketGuildUser)command.User, botConfig)) { await command.FollowupAsync("❌ 權限不足，此指令僅限管理員使用。"); return; }

                            string dateStr = (string)command.Data.Options.First().Value;
                            if (!DateTime.TryParse(dateStr, out DateTime dt)) { await command.FollowupAsync("❌ 日期格式錯誤。"); break; }
                            var dLogs = await db.PointLogs.Where(l => l.GuildId == gid && l.Timestamp.Date == dt.Date && !l.IsDeleted).ToListAsync();
                            if (!dLogs.Any()) { await command.FollowupAsync("📭 該日無紀錄。"); break; }

                            var dSb = new StringBuilder($"#  {dt:yyyy-MM-dd} 點數紀錄清單\n```text\n");
                            foreach (var l in dLogs)
                            {
                                string idStr = $"[ID: {l.Id}]".PadRight(9);
                                string timeStr = l.Timestamp.ToString("HH:mm");
                                string nameStr = l.RobloxUsername.PadRight(18);
                                // 自動判斷正負號
                                string ptsStr = (l.PointsAdded > 0 ? $"+{l.PointsAdded}" : l.PointsAdded.ToString()).PadLeft(5);
                                string adminStr = l.AdminName.PadRight(12);

                                dSb.AppendLine($"{idStr} {timeStr} | {nameStr} ➔ {ptsStr} 點 | 登記: {adminStr} | 原因: {l.Reason}");
                            }
                            dSb.AppendLine("```");
                            await command.FollowupAsync(dSb.ToString());
                            break;
                        }
                    case "points":
                        {
                            var targetUser = (SocketGuildUser)command.Data.Options.First(x => x.Name == "user").Value;
                            string targetName = targetUser.Nickname ?? targetUser.Username;
                            if (targetName.Contains("]")) targetName = targetName.Substring(targetName.LastIndexOf(']') + 1).Trim();
                            var userPoint = await db.UserPoints.FirstOrDefaultAsync(u => u.GuildId == gid && u.RobloxUsername.ToLower() == targetName.ToLower());
                            int currentPoints = userPoint != null ? userPoint.Points : 0;
                            string fetchTime = Program.GetTaipeiTime().ToString("yyyy-MM-dd HH:mm:ss");
                            await command.FollowupAsync($" **{targetName}** 目前擁有 **{currentPoints}** 點。\n>  *數據抓取時間：{fetchTime}*");
                            break;
                        }

                    case "admin-setup":
                        {
                            // 你的專屬開發者 ID
                            ulong developerId = 1018018187318673471;

                            // 只有 Discord 最高管理員「或是」你本人可以設定這個頻道
                            if (!((SocketGuildUser)command.User).GuildPermissions.Administrator && command.User.Id != developerId)
                            {
                                await command.FollowupAsync("❌ 權限不足！只有本伺服器的「Discord 伺服器管理員」或「系統開發者」可以進行初步設定。");
                                return;
                            }

                            string codes = ((string)command.Data.Options.First(x => x.Name == "server_codes").Value).ToUpper();
                            var mndRole = (SocketRole)command.Data.Options.First(x => x.Name == "mnd_role").Value;
                            // 👇 新增這行：嘗試讀取第二個身分組
                            var mndRole2 = (SocketRole)command.Data.Options.FirstOrDefault(x => x.Name == "mnd_role_2")?.Value;
                            var adminChannel = await db.AdminChannels.FindAsync(command.Channel.Id);

                            if (adminChannel == null)
                            {
                                // 👇 修改這行：將 MndAdminRoleId2 也存進去
                                db.AdminChannels.Add(new AdminChannelConfig { ChannelId = command.Channel.Id, GuildId = gid, TargetServerCodes = codes, MndAdminRoleId = mndRole.Id, MndAdminRoleId2 = mndRole2?.Id });
                            }
                            else
                            {
                                adminChannel.TargetServerCodes = codes;
                                adminChannel.MndAdminRoleId = mndRole.Id;
                                adminChannel.MndAdminRoleId2 = mndRole2?.Id; // 👇 新增這行
                            }
                            // 👇 新增判斷：為了讓成功訊息正確顯示有幾個身分組
                            string roleMsg = $"<@&{mndRole.Id}>";
                            if (mndRole2 != null) roleMsg += $" 及 <@&{mndRole2.Id}>";
                            await db.SaveChangesAsync();
                            await command.FollowupAsync($"✅ **已將此頻道設為國防部推播頻道！**\n📡 授權監控的單位編號：`{codes}`\n👮 授權查詢身分組：<@&{mndRole.Id}>\n> *(擁有該身分組的長官，現在可以在此頻道使用 `/admin-view` 指令查詢各單位排行榜了！)*");
                            break;
                        }
                    case "add-menu-item":
                        {
                            if (botConfig == null) { await command.FollowupAsync("❌ 請先設定機器人。"); return; }
                            if (!IsAdmin((SocketGuildUser)command.User, botConfig)) { await command.FollowupAsync("❌ 限管理員執行。"); return; }

                            string itemName = (string)command.Data.Options.First(x => x.Name == "item_name").Value;
                            int points = Convert.ToInt32((long)command.Data.Options.First(x => x.Name == "points").Value);

                            string noteType = command.Data.Options.FirstOrDefault(x => x.Name == "note_type")?.Value as string ?? "備註";
                            string noteText = command.Data.Options.FirstOrDefault(x => x.Name == "note")?.Value as string;

                            // 👇 讀取使用者輸入的插隊目標
                            string insertAfter = command.Data.Options.FirstOrDefault(x => x.Name == "insert_after")?.Value as string;

                            string finalNote = null;
                            if (!string.IsNullOrEmpty(noteText))
                            {
                                finalNote = $"{noteType}: {noteText}";
                            }

                            // 如果沒有指定插隊，就按照原本的方式新增到最後面
                            if (string.IsNullOrEmpty(insertAfter))
                            {
                                db.RewardMenuItems.Add(new RewardMenuItem { GuildId = gid, ItemName = itemName, Points = points, Note = finalNote });
                                await db.SaveChangesAsync();

                                string msg = points <= 0 ? $"✅ 已新增分類標題：`{itemName}`" : $"✅ 已新增品項：`{itemName}` (點數: {points})";
                                await command.FollowupAsync($"{msg}\n> 此更動目前僅存在於草稿中。請使用 `/edit-menu` 預覽並申請發布。");
                            }
                            else
                            {
                                // 有指定插隊目標
                                var existingItems = await db.RewardMenuItems.Where(x => x.GuildId == gid).OrderBy(x => x.Id).ToListAsync();
                                var targetItem = existingItems.FirstOrDefault(x => x.ItemName == insertAfter);

                                if (targetItem == null)
                                {
                                    await command.FollowupAsync($"❌ 草稿中找不到名稱為 `{insertAfter}` 的品項。請確認輸入的名稱是否完全相符。");
                                    return;
                                }

                                // 先將新品項加到資料庫取得 ID
                                var newItem = new RewardMenuItem { GuildId = gid, ItemName = itemName, Points = points, Note = finalNote };
                                db.RewardMenuItems.Add(newItem);
                                await db.SaveChangesAsync(); // 儲存後 newItem 會獲得最新的 Id，暫時排在最後面

                                // 重新抓取包含新品項的完整清單 (依 Id 排序)
                                existingItems = await db.RewardMenuItems.Where(x => x.GuildId == gid).OrderBy(x => x.Id).ToListAsync();

                                // 建立我們「期望的排序資料」快照 (使用 Tuple 避免物件參考覆蓋的問題)
                                var desiredData = new List<(string Name, int Points, string Note)>();
                                foreach (var item in existingItems.Where(i => i.Id != newItem.Id))
                                {
                                    desiredData.Add((item.ItemName, item.Points, item.Note));
                                    // 找到目標品項後，緊接著把新品項塞進期望的清單中
                                    if (item.ItemName == insertAfter)
                                    {
                                        desiredData.Add((newItem.ItemName, newItem.Points, newItem.Note));
                                    }
                                }

                                // 將現有資料庫列的內容，依序替換成期望的內容 (完美實現平移插隊)
                                for (int i = 0; i < existingItems.Count; i++)
                                {
                                    existingItems[i].ItemName = desiredData[i].Name;
                                    existingItems[i].Points = desiredData[i].Points;
                                    existingItems[i].Note = desiredData[i].Note;
                                }

                                await db.SaveChangesAsync();

                                string msg = points <= 0 ? $"✅ 已成功將分類標題 `{itemName}` 安插在 `{insertAfter}` 之後！" : $"✅ 已成功將品項 `{itemName}` (點數: {points}) 安插在 `{insertAfter}` 之後！";
                                await command.FollowupAsync($"{msg}\n> 此更動目前僅存在於草稿中。請使用 `/edit-menu` 預覽並發布。");
                            }
                            break;
                        }

                    case "remove-menu-item":
                        {
                            if (botConfig == null) { await command.FollowupAsync("❌ 請先設定機器人。"); return; }
                            if (!IsAdmin((SocketGuildUser)command.User, botConfig)) { await command.FollowupAsync("❌ 限管理員執行。"); return; }

                            string itemName = (string)command.Data.Options.First(x => x.Name == "item_name").Value;
                            var items = await db.RewardMenuItems.Where(x => x.GuildId == gid && x.ItemName == itemName).ToListAsync();

                            if (!items.Any()) { await command.FollowupAsync($"❌ 草稿中找不到名稱為 `{itemName}` 的品項。"); return; }

                            db.RewardMenuItems.RemoveRange(items);
                            await db.SaveChangesAsync();
                            await command.FollowupAsync($"✅ 已從草稿中移除品項：`{itemName}`\n> 請使用 `/edit-menu` 確認最新狀態。");
                            break;
                        }
                    case "change-menu-item":
                        {
                            if (botConfig == null) { await command.FollowupAsync("❌ 請先設定機器人。"); return; }
                            if (!IsAdmin((SocketGuildUser)command.User, botConfig)) { await command.FollowupAsync("❌ 限管理員執行。"); return; }

                            string oldItemName = (string)command.Data.Options.First(x => x.Name == "old_item_name").Value;
                            string newItemName = (string)command.Data.Options.First(x => x.Name == "new_item_name").Value;
                            int points = Convert.ToInt32((long)command.Data.Options.First(x => x.Name == "points").Value);

                            // 👇 讀取標籤類型與內容
                            string noteType = command.Data.Options.FirstOrDefault(x => x.Name == "note_type")?.Value as string ?? "備註";
                            string noteText = command.Data.Options.FirstOrDefault(x => x.Name == "note")?.Value as string;

                            string finalNote = null;
                            if (!string.IsNullOrEmpty(noteText))
                            {
                                finalNote = $"{noteType}: {noteText}";
                            }

                            // 尋找要被替換的舊品項
                            var itemToEdit = await db.RewardMenuItems.FirstOrDefaultAsync(x => x.GuildId == gid && x.ItemName == oldItemName);

                            if (itemToEdit == null)
                            {
                                await command.FollowupAsync($"❌ 草稿中找不到名稱為 `{oldItemName}` 的品項。請確認輸入的名稱是否完全相符。");
                                return;
                            }

                            // 覆寫內容
                            itemToEdit.ItemName = newItemName;
                            itemToEdit.Points = points;
                            itemToEdit.Note = finalNote; // 👇 存入包含標籤的內容

                            await db.SaveChangesAsync();

                            string msg = points <= 0 ? $"✅ 已成功替換為分類標題：`{newItemName}`" : $"✅ 已成功替換為品項：`{newItemName}` (點數: {points})";
                            await command.FollowupAsync($"{msg}\n> *(原排列順序已保留)* 此更動目前僅存在於草稿中，請使用 `/edit-menu` 預覽並發布。");
                            break;
                        }


                    case "menu":
                        {
                            if (botConfig == null || string.IsNullOrEmpty(botConfig.RewardMenuContent))
                            {
                                await command.FollowupAsync("📭 目前尚未發布任何 MENU，請管理員使用 `/edit-menu` 進行發布。");
                                break;
                            }

                            string timeStr = botConfig.RewardMenuUpdateTime?.ToString("yyyy年MM月dd日HH:mm") ?? "未知時間";
                            // 🚨 這裡不需要再 Format，因為我們發布時就會直接把排版好的 ANSI 存進資料庫
                            string msg = $"### 🌟 點數兌換 MENU 總覽\n\n{botConfig.RewardMenuContent}\n\n> *此為 {timeStr} 公告*";
                            await command.FollowupAsync(msg);
                            break;
                        }

                    case "view-admins":
                        {
                            if (botConfig == null) { await command.FollowupAsync("❌ 未設定。"); break; }
                            if (!IsAdmin((SocketGuildUser)command.User, botConfig)) { await command.FollowupAsync("❌ 權限不足，僅限管理員查看。"); return; }

                            var roles = botConfig.AdminRoleIds?.Split(',').Where(s => !string.IsNullOrEmpty(s)).ToList() ?? new List<string>();
                            if (!roles.Any()) roles.Add(botConfig.AdminRoleId.ToString()); // 相容舊版資料

                            var roleMentions = roles.Select(id => $"<@&{id}>");
                            await command.FollowupAsync($"👮 **目前的管理員身分組名單：**\n{string.Join("\n", roleMentions)}");
                            break;
                        }
                    case "bind-sheet":
                        {
                            if (botConfig == null || string.IsNullOrEmpty(botConfig.RobloxGroupId))
                            {
                                await command.FollowupAsync("❌ 請先使用 `/setup-roca` 完成基本設定。");
                                break;
                            }

                            // 判斷該伺服器所屬單位
                            string sheetId = "";
                            string unitName = "";
                            string combinedSheetId = "1v8TmW04kJwUKPFeghK5OkNzyj65XSrcHXHWGfsokTQ8";

                            switch (botConfig.RobloxGroupId)
                            {
                                case "13549943":
                                    unitName = "憲兵";
                                    sheetId = "1cSaJJXq75MJ479RV5KcZ8oYBHxdbexC5f0XQtzd7Bwo";
                                    break;
                                case "13662982":
                                    unitName = "裝甲";
                                    sheetId = "1k08y_tUTA0_eS_W9zkGQzHU3OSLHcVlzKtTXvjzUxB4";
                                    break;
                                case "16223475":
                                    unitName = "航特";
                                    sheetId = "1rOXjrRdBD1S77guAByBaACgxToReHMUhSiU2WE1_zrE";
                                    break;
                                default:
                                    sheetId = botConfig.GoogleSheetId;
                                    unitName = "自訂單位";
                                    break;
                            }

                            if (string.IsNullOrEmpty(sheetId))
                            {
                                await command.FollowupAsync("❌ 目前無法辨識您的單位，或者此伺服器尚未綁定自訂試算表。");
                                break;
                            }

                            // 建立前往試算表的網址
                            string sheetUrl = $"https://docs.google.com/spreadsheets/d/{sheetId}/edit";
                            string combinedUrl = $"https://docs.google.com/spreadsheets/d/{combinedSheetId}/edit";

                            // 👇 判斷執行指令的使用者是否為管理員
                            bool isAdmin = IsAdmin((SocketGuildUser)command.User, botConfig);

                            // 建立按鈕介面 (先加入所有人都能看的單一單位試算表)
                            var buttons = new ComponentBuilder()
                                .WithButton($"🔗 開啟 {unitName} 專屬試算表", style: ButtonStyle.Link, url: sheetUrl);

                            // 👇 如果是管理員，才額外加入總覽表與強制同步按鈕
                            if (isAdmin)
                            {
                                buttons.WithButton("🔗 開啟 三單位總覽試算表", style: ButtonStyle.Link, url: combinedUrl)
                                       .WithButton("🔄 強制同步最新資料至試算表", $"force_sync_sheet_{gid}", ButtonStyle.Success);
                            }

                            // 根據身分顯示不同的文案
                            string replyMessage = isAdmin
                                ? $"✅ **系統已自動對應您的單位為：【{unitName}】**\n> 點擊下方按鈕即可快速前往您的資料庫，或手動執行資料同步："
                                : $"✅ **您的所屬單位為：【{unitName}】**\n> 點擊下方按鈕即可查看本單位的專屬排行榜：";

                            await command.FollowupAsync(replyMessage, components: buttons.Build());
                            break;
                        }
                    case "admin-view":
                        {
                            var ac = await db.AdminChannels.FindAsync(command.Channel.Id);
                            if (ac == null) { await command.FollowupAsync("❌ 此頻道尚未被設定為 Admin 監控頻道。"); return; }

                            var execUser = (SocketGuildUser)command.User;

                            // 👇 修改這裡：分別檢查兩種身分組
                            bool hasMndRole1 = ac.MndAdminRoleId.HasValue && execUser.Roles.Any(r => r.Id == ac.MndAdminRoleId.Value);
                            bool hasMndRole2 = ac.MndAdminRoleId2.HasValue && execUser.Roles.Any(r => r.Id == ac.MndAdminRoleId2.Value);

                            // 👇 修改這行：只要有其中一種身分組，或是伺服器管理員，就放行
                            if (!execUser.GuildPermissions.Administrator && !hasMndRole1 && !hasMndRole2)
                            {
                                await command.FollowupAsync("❌ 權限不足！您沒有被授權使用總部查詢指令。");
                                return;
                            }

                            string targetCode = ((string)command.Data.Options.First().Value).ToUpper();
                            if (!ac.TargetServerCodes.Contains(targetCode)) { await command.FollowupAsync($"❌ 此頻道未被授權查看編號 `{targetCode}` 的資料庫。"); return; }

                            var targetConfig = await db.Configs.FirstOrDefaultAsync(c => c.ServerCode == targetCode);
                            if (targetConfig == null) { await command.FollowupAsync("❌ 找不到該編號的伺服器，請確認單位是否已綁定機器人。"); return; }

                            var targetData = await db.UserPoints.Where(u => u.GuildId == targetConfig.GuildId).OrderByDescending(u => u.Points).ToListAsync();
                            if (!targetData.Any()) { await command.FollowupAsync($"📭 目標編號 `{targetCode}` 尚無成員資料。"); break; }

                            var targetSb = new StringBuilder($"## 國防部查閱：單位 [{targetCode}] 點數總覽\n```ansi\n");
                            int tRank = 1;
                            foreach (var u in targetData)
                            {
                                string tRankStr = $"\u001b[34m[{tRank}]\u001b[0m".PadRight(14);
                                string tPointStr = $"\u001b[32m{u.Points} 點\u001b[0m".PadLeft(17);
                                targetSb.AppendLine($"{tRankStr} {u.RobloxUsername.PadRight(20)} ➔ {tPointStr}");
                                tRank++;
                                if (targetSb.Length > 1800) break;
                            }
                            targetSb.AppendLine("```");
                            await command.FollowupAsync(targetSb.ToString());
                            break;
                        }
                }

            }
            catch (Exception ex)
            {
                string realError = ex.InnerException != null ? ex.InnerException.Message : ex.Message;
                Console.WriteLine($"[指令錯誤] {realError}");
                string errMsg = $"❌ 發生內部錯誤: `{realError}`\n請檢查資料庫狀態。";
                // 👇 新增這行：發生例外錯誤時自動私訊回報給你
                _ = SendStatusToDeveloperAsync($"🚨 **[錯誤回報]** 發生未預期的例外錯誤！\n> **觸發指令**：`{command.Data.Name}`\n> **錯誤內容**：`{realError}`\n> **時間**：{Program.GetTaipeiTime():yyyy-MM-dd HH:mm:ss}");
                if (command.HasResponded) await command.FollowupAsync(errMsg, ephemeral: true);
                else await command.RespondAsync(errMsg, ephemeral: true);
            }
        }
        // 👇 生成與更新結算面板 UI
        private async Task UpdateEventDashboardAsync(IDiscordInteraction interaction, string sessId, SocketGuild guild, bool isFirst = false)
        {
            if (!_pendingDistributions.TryGetValue(sessId, out var pending)) return;

            // 👇 計算總時長
            var duration = pending.EndTime - pending.StartTime;
            string durationStr = $"{(int)duration.TotalHours} 小時 {duration.Minutes} 分鐘 {duration.Seconds} 秒";

            var sb = new StringBuilder();
            sb.AppendLine($"## 📊 語音活動結算面板 (審核中)");
            sb.AppendLine($"> **發放原因：** {pending.Reason}");
            sb.AppendLine($"> **活動總時長：** `{durationStr}`"); // 👈 顯示在面板上
            sb.AppendLine($"> **預定點數：** `{pending.Points}` 點");
            sb.AppendLine($"> **符合資格：** `{pending.FinalUserIds.Count}` 人 *(原本條件為待滿5分鐘)*\n");

            sb.AppendLine("👥 **【最終核定發放名單】**");
            if (pending.FinalUserIds.Any())
                sb.AppendLine(string.Join(", ", pending.FinalUserIds.Select(id => $"<@{id}>")));
            else
                sb.AppendLine("*無人符合資格，或名單已被清空。*");

            sb.AppendLine("\n📜 **【進出入紀錄 (節錄最新)】**\n```ansi");
            var logsToDisplay = pending.ActionLogs.Skip(Math.Max(0, pending.ActionLogs.Count - 15)).ToList();
            if (pending.ActionLogs.Count > 15) sb.AppendLine($"... (省略前方 {pending.ActionLogs.Count - 15} 筆)");
            foreach (var log in logsToDisplay) sb.AppendLine(log);
            sb.AppendLine("```");

            var cb = new ComponentBuilder()
                .WithSelectMenu(new SelectMenuBuilder().WithCustomId($"evadd_{sessId}").WithPlaceholder("➕ 選擇破例補登的人員 (可複選)").WithType(ComponentType.UserSelect).WithMaxValues(25))
                .WithSelectMenu(new SelectMenuBuilder().WithCustomId($"evrem_{sessId}").WithPlaceholder("➖ 選擇要剔除的人員 (可複選)").WithType(ComponentType.UserSelect).WithMaxValues(25))
                .WithButton("✅ 確定名單並一鍵發放", $"evok_{sessId}", ButtonStyle.Success)
                .WithButton("❌ 取消作廢", $"evcancel_{sessId}", ButtonStyle.Danger);

            if (isFirst && interaction is SocketSlashCommand cmd) await cmd.FollowupAsync(sb.ToString(), components: cb.Build());
            else if (interaction is SocketMessageComponent comp) await comp.UpdateAsync(m => { m.Content = sb.ToString(); m.Components = cb.Build(); });
        }
        // 自動將資料庫的品項清單轉換為精美排版與 ANSI 色塊
        // 自動將管理員的純文字轉換為精美排版與 ANSI 色塊
        private string FormatMenuItems(List<RewardMenuItem> items)
        {
            if (!items.Any()) return "```ansi\n\u001b[33m目前草稿中尚無任何品項\u001b[0m\n```";

            var sb = new StringBuilder();
            sb.AppendLine("```ansi");

            foreach (var item in items)
            {
                if (item.Points <= 0)
                {
                    sb.AppendLine($"\u001b[33m-- {item.ItemName} --\u001b[0m");
                }
                else
                {
                    string paddedItem = PadRightCustom($"[{item.ItemName}]", 18);
                    string itemStr = $"\u001b[34m{paddedItem}\u001b[0m"; // 藍色

                    string paddedPts = item.Points.ToString().PadLeft(4);
                    string ptsStr = $"\u001b[32m{paddedPts} 點\u001b[0m"; // 綠色

                    string formattedLine = $"{itemStr} ➔ {ptsStr}";

                    if (!string.IsNullOrEmpty(item.Note))
                    {
                        // 👇 判斷如果字串已經包含了「條件:」或「備註:」，就直接顯示；如果沒有(舊資料)，就加上預設的「條件/備註:」
                        if (item.Note.StartsWith("條件:") || item.Note.StartsWith("備註:"))
                        {
                            formattedLine += $" \u001b[30m| {item.Note}\u001b[0m";
                        }
                        else
                        {
                            formattedLine += $" \u001b[30m| 條件/備註: {item.Note}\u001b[0m";
                        }
                    }
                    sb.AppendLine(formattedLine);
                }
            }
            sb.AppendLine("```");
            return sb.ToString();
        }


        private async Task HandleInteractionAsync(SocketInteraction interaction)
        {
            if (interaction is SocketMessageComponent component)
            {
                string id = component.Data.CustomId;
                ulong gid = (ulong)component.GuildId;
                var executor = (SocketGuildUser)component.User;

                using var db = new BotDbContext(_configuration);
                var botConfig = await db.Configs.FindAsync(gid);

                // 檢查點擊按鈕的人是否具備管理員權限
                bool isAdmin = IsAdmin(executor, botConfig);

                if (id.StartsWith("clear_cancel_"))
                {
                    if (!isAdmin) { await component.RespondAsync("❌ 只有管理員可以取消此操作。", ephemeral: true); return; }
                    await component.UpdateAsync(m => { m.Content = "❌ 資料清空操作已由管理員取消。"; m.Components = null; });
                    return;
                }

                if (id.StartsWith("clear_step1_"))
                {
                    if (!isAdmin) { await component.RespondAsync("❌ 權限不足，僅限管理員確認。", ephemeral: true); return; }

                    // 將第一位管理員的 ID 記錄在按鈕的 CustomId 中，傳給下一步
                    var btn2 = new ComponentBuilder()
                        .WithButton("🚨 第二位管理員最終確認", $"clear_step2_{gid}_{executor.Id}", ButtonStyle.Danger)
                        .WithButton("取消", $"clear_cancel_{gid}", ButtonStyle.Secondary);

                    await component.UpdateAsync(m => {
                        m.Content = $"🚨 **【最終警告】** 資料將永久刪除！\n> 第一位管理員 {executor.Mention} 已確認。\n> 需要 **第二位不同的管理員** 按下確認才能執行！";
                        m.Components = btn2.Build();
                    });

                    // 額外發送一則新訊息來真正觸發 Ping 推播通知
                    if (botConfig != null)
                    {
                        string adminRoleMention = $"<@&{botConfig.AdminRoleId}>";
                        await component.Channel.SendMessageAsync($"🔔 {adminRoleMention} 警告：{executor.Mention} 正在申請清空資料庫，請另一位管理員至上方訊息進行最終確認！");
                    }
                    return;
                }

                if (id.StartsWith("clear_step2_"))
                {
                    if (!isAdmin) { await component.RespondAsync("❌ 權限不足，僅限管理員確認。", ephemeral: true); return; }

                    // 解析出第一位管理員的 ID
                    var parts = id.Split('_');
                    if (parts.Length >= 4 && ulong.TryParse(parts[3], out ulong firstAdminId))
                    {
                        // 防呆機制：同一個人不能按兩次
                        if (executor.Id == firstAdminId)
                        {
                            await component.RespondAsync("❌ 您已經確認過了！必須由 **另一位不同的管理員** 來進行最終確認。", ephemeral: true);
                            return;
                        }

                        // 1. 將所有人的點數歸零，但不刪除資料庫內的人員名單
                        var allUsers = db.UserPoints.Where(u => u.GuildId == gid);
                        foreach (var user in allUsers)
                        {
                            user.Points = 0;
                        }

                        // 2. 將先前的歷史紀錄標記為已刪除
                        foreach (var log in db.PointLogs.Where(l => l.GuildId == gid))
                        {
                            log.IsDeleted = true;
                        }

                        await db.SaveChangesAsync();

                        // 3. 更新原本的按鈕訊息
                        await component.UpdateAsync(m => {
                            m.Content = $"🔥 **全體人員點數已成功歸零！**\n> 授權執行者：<@{firstAdminId}> 與 {executor.Mention}。";
                            m.Components = null;
                        });

                        // 4. 👇 新增：抓取第一位管理員的名稱，並發送推播通知給國防部 Admin 頻道 👇
                        string firstAdminName = executor.Guild.GetUser(firstAdminId)?.Username ?? firstAdminId.ToString();
                        string secondAdminName = executor.Username;

                        _ = BroadcastToAdminChannelsAsync(gid,
                            $"🚨 **【重大操作警告：全體點數歸零】**\n" +
                            $"> 該單位的負責人已執行了 **清空所有成員點數** 的操作！\n" +
                            $"> 授權執行者一：**{firstAdminName}**\n" +
                            $"> 授權執行者二：**{secondAdminName}**\n" +
                            $"> 狀態：資料庫點數已全數歸零。", true); // 👈 這裡加上 true
                        // 👆 ========================================================== 👆
                    }
                    else
                    {
                        await component.RespondAsync("❌ 發生內部錯誤，無法辨識第一位管理員的身分。", ephemeral: true);
                    }
                }
                // ==================== 解除綁定 (Unbind) 雙重確認邏輯 ====================

                // 處理取消按鈕
                if (id.StartsWith("unbind_cancel_"))
                {
                    if (!isAdmin) { await component.RespondAsync("❌ 只有管理員可以取消此操作。", ephemeral: true); return; }
                    await component.UpdateAsync(m => { m.Content = "❌ 解除綁定操作已由管理員取消。"; m.Components = null; });
                    return;
                }

                // 處理第一位管理員確認
                if (id.StartsWith("unbind_step1_"))
                {
                    if (!isAdmin) { await component.RespondAsync("❌ 權限不足，僅限管理員確認。", ephemeral: true); return; }

                    // 將第一位管理員的 ID 記錄在按鈕的 CustomId 中，傳給下一步
                    var btn2 = new ComponentBuilder()
                        .WithButton("🚨 第二位管理員最終確認", $"unbind_step2_{gid}_{executor.Id}", ButtonStyle.Danger)
                        .WithButton("取消", $"unbind_cancel_{gid}", ButtonStyle.Secondary);

                    await component.UpdateAsync(m => {
                        m.Content = $"🚨 **【最終警告】** 伺服器綁定設定將被解除！\n> 第一位管理員 {executor.Mention} 已確認。\n> 需要 **第二位不同的管理員** 按下確認才能執行！";
                        m.Components = btn2.Build();
                    });

                    // 推播提醒管理員身分組
                    if (botConfig != null)
                    {
                        string adminRoleMention = $"<@&{botConfig.AdminRoleId}>";
                        await component.Channel.SendMessageAsync($"🔔 {adminRoleMention} 警告：{executor.Mention} 正在申請解除伺服器綁定，請另一位管理員至上方訊息進行最終確認！");
                    }
                    return;
                }

                // 處理第二位管理員最終確認
                if (id.StartsWith("unbind_step2_"))
                {
                    if (!isAdmin) { await component.RespondAsync("❌ 權限不足，僅限管理員確認。", ephemeral: true); return; }

                    // 解析出第一位管理員的 ID
                    var parts = id.Split('_');
                    if (parts.Length >= 4 && ulong.TryParse(parts[3], out ulong firstAdminId))
                    {
                        // 防呆機制：同一個人不能按兩次
                        if (executor.Id == firstAdminId)
                        {
                            await component.RespondAsync("❌ 您已經確認過了！必須由 **另一位不同的管理員** 來進行最終確認。", ephemeral: true);
                            return;
                        }

                        // 1. 執行解除綁定邏輯
                        if (botConfig != null)
                        {
                            db.Configs.Remove(botConfig);
                            await db.SaveChangesAsync();
                        }

                        // 2. 更新原本的按鈕訊息
                        await component.UpdateAsync(m => {
                            m.Content = $"🔓 **本伺服器已成功解除綁定！**\n> 授權執行者：<@{firstAdminId}> 與 {executor.Mention}。";
                            m.Components = null;
                        });

                        // 3. (選擇性) 推播通知給國防部 Admin 頻道
                        string firstAdminName = executor.Guild.GetUser(firstAdminId)?.Username ?? firstAdminId.ToString();
                        string secondAdminName = executor.Username;

                        _ = BroadcastToAdminChannelsAsync(gid,
                            $"🚨 **【重大操作警告：解除綁定】**\n" +
                            $"> 該單位已解除了 Discord 伺服器綁定！\n" +
                            $"> 授權執行者一：**{firstAdminName}**\n" +
                            $"> 授權執行者二：**{secondAdminName}**", true);
                    }
                    else
                    {
                        await component.RespondAsync("❌ 發生內部錯誤，無法辨識第一位管理員的身分。", ephemeral: true);
                    }
                    return;
                }
                // =======================================================================
                // ==================== 變更管理員身分組 (Edit Admin) 雙重確認邏輯 ====================

                // 處理取消按鈕
                if (id.StartsWith("editadm_cancel_"))
                {
                    if (!isAdmin) { await component.RespondAsync("❌ 只有管理員可以取消此操作。", ephemeral: true); return; }
                    await component.UpdateAsync(m => { m.Content = "❌ 權限變更操作已由管理員取消。"; m.Components = null; });
                    return;
                }

                // 處理第一位管理員確認
                if (id.StartsWith("editadm_s1_"))
                {
                    if (!isAdmin) { await component.RespondAsync("❌ 權限不足，僅限管理員確認。", ephemeral: true); return; }

                    var parts = id.Split('_');
                    string action = parts[3];
                    ulong roleId = ulong.Parse(parts[4]);
                    string actionText = action == "add" ? "新增" : "移除";

                    // 夾帶第一位管理員 ID
                    var btn2 = new ComponentBuilder()
                        .WithButton("🚨 第二位管理員最終確認", $"editadm_s2_{gid}_{action}_{roleId}_{executor.Id}", ButtonStyle.Danger)
                        .WithButton("取消", $"editadm_cancel_{gid}", ButtonStyle.Secondary);

                    await component.UpdateAsync(m => {
                        m.Content = $"🚨 **【最終警告】** 即將 **{actionText}** <@&{roleId}> 的管理員權限！\n> 第一位管理員 {executor.Mention} 已確認。\n> 需要 **第二位不同的管理員** 按下確認才能執行！";
                        m.Components = btn2.Build();
                    });

                    if (botConfig != null)
                    {
                        string adminRoleMention = $"<@&{botConfig.AdminRoleId}>";
                        await component.Channel.SendMessageAsync($"🔔 {adminRoleMention} 警告：{executor.Mention} 正在申請 {actionText} 管理身分組，請另一位管理員至上方訊息進行最終確認！");
                    }
                    return;
                }

                // 處理第二位管理員最終確認
                if (id.StartsWith("editadm_s2_"))
                {
                    if (!isAdmin) { await component.RespondAsync("❌ 權限不足，僅限管理員確認。", ephemeral: true); return; }

                    var parts = id.Split('_');
                    string action = parts[3];
                    ulong roleId = ulong.Parse(parts[4]);
                    ulong firstAdminId = ulong.Parse(parts[5]);

                    if (executor.Id == firstAdminId)
                    {
                        await component.RespondAsync("❌ 您已經確認過了！必須由 **另一位不同的管理員** 來進行最終確認。", ephemeral: true);
                        return;
                    }

                    if (botConfig != null)
                    {
                        var currentRoles = botConfig.AdminRoleIds?.Split(',').Where(s => !string.IsNullOrEmpty(s)).Select(ulong.Parse).ToList() ?? new List<ulong>();

                        if (action == "add") { if (!currentRoles.Contains(roleId)) currentRoles.Add(roleId); }
                        else if (action == "remove") { if (currentRoles.Contains(roleId)) currentRoles.Remove(roleId); }

                        // 重新組合成字串並更新
                        botConfig.AdminRoleIds = string.Join(",", currentRoles);
                        if (currentRoles.Any()) botConfig.AdminRoleId = currentRoles.First(); // 同步更新舊版相容的單一身分組 ID

                        await db.SaveChangesAsync();

                        string actionText = action == "add" ? "新增" : "移除";
                        await component.UpdateAsync(m => {
                            m.Content = $"✅ **權限變更成功！** 已成功 **{actionText}** <@&{roleId}> 的管理員權限。\n> 授權執行者：<@{firstAdminId}> 與 {executor.Mention}。";
                            m.Components = null;
                        });
                        // 【將這行加進 editadm_s2_ 處理成功的最後面】
                        _ = BroadcastToAdminChannelsAsync(gid, $"🚨 **【管理員權限變更】**\n> 動作：**{actionText}** <@&{roleId}>\n> 授權執行者：<@{firstAdminId}> 與 {executor.Mention}", true);
                    }
                    return;
                }
                // ==================== MENU 預覽發布 雙重確認邏輯 ====================
                if (id.StartsWith("menuno_"))
                {
                    if (!isAdmin) { await component.RespondAsync("❌ 權限不足。", ephemeral: true); return; }
                    await component.UpdateAsync(m => { m.Content = "❌ 已取消發布申請，您可以繼續使用 `/add-menu-item` 或 `/remove-menu-item` 修改草稿。"; m.Components = null; });
                    return;
                }
                // 1. 單人直接清空資料
                if (id.StartsWith("clear_single_"))
                {
                    if (!isAdmin) { await component.RespondAsync("❌ 權限不足。", ephemeral: true); return; }
                    foreach (var user in db.UserPoints.Where(u => u.GuildId == gid)) user.Points = 0;
                    foreach (var log in db.PointLogs.Where(l => l.GuildId == gid)) log.IsDeleted = true;
                    await db.SaveChangesAsync();
                    await component.UpdateAsync(m => { m.Content = $"🔥 **全體人員點數已成功歸零！**\n> (單人確認模式) 執行者：{executor.Mention}。"; m.Components = null; });
                    _ = BroadcastToAdminChannelsAsync(gid, $"🚨 **【重大操作警告：全體點數歸零】**\n> 執行者：**{executor.Username}** (單人確認模式)", true);
                    return;
                }

                // 2. 單人直接解除綁定
                if (id.StartsWith("unbind_single_"))
                {
                    if (!isAdmin) { await component.RespondAsync("❌ 權限不足。", ephemeral: true); return; }
                    if (botConfig != null) { db.Configs.Remove(botConfig); await db.SaveChangesAsync(); }
                    await component.UpdateAsync(m => { m.Content = $"🔓 **本伺服器已成功解除綁定！**\n> (單人確認模式) 執行者：{executor.Mention}。"; m.Components = null; });
                    _ = BroadcastToAdminChannelsAsync(gid, $"🚨 **【重大操作警告：解除綁定】**\n> 執行者：**{executor.Username}** (單人確認模式)", true);
                    return;
                }

                // 3. 單人直接變更管理員
                if (id.StartsWith("editadm_single_"))
                {
                    if (!isAdmin) { await component.RespondAsync("❌ 權限不足。", ephemeral: true); return; }
                    var parts = id.Split('_'); string action = parts[3]; ulong roleId = ulong.Parse(parts[4]);
                    if (botConfig != null)
                    {
                        var currentRoles = botConfig.AdminRoleIds?.Split(',').Where(s => !string.IsNullOrEmpty(s)).Select(ulong.Parse).ToList() ?? new List<ulong>();
                        if (action == "add") { if (!currentRoles.Contains(roleId)) currentRoles.Add(roleId); }
                        else if (action == "remove") { if (currentRoles.Contains(roleId)) currentRoles.Remove(roleId); }
                        botConfig.AdminRoleIds = string.Join(",", currentRoles);
                        if (currentRoles.Any()) botConfig.AdminRoleId = currentRoles.First();
                        await db.SaveChangesAsync();
                        string actionText = action == "add" ? "新增" : "移除";
                        await component.UpdateAsync(m => { m.Content = $"✅ **權限變更成功！** 已成功 **{actionText}** <@&{roleId}> 的管理員權限。\n> (單人確認模式) 執行者：{executor.Mention}。"; m.Components = null; });
                        _ = BroadcastToAdminChannelsAsync(gid, $"🚨 **【管理員權限變更】**\n> 動作：**{actionText}** <@&{roleId}>\n> 執行者：{executor.Mention} (單人確認模式)", true);
                    }
                    return;
                }

                // 4. 新的單人發布 MENU 邏輯
                if (id.StartsWith("menuok_single_"))
                {
                    if (!isAdmin) { await component.RespondAsync("❌ 權限不足。", ephemeral: true); return; }
                    var items = await db.RewardMenuItems.Where(x => x.GuildId == gid).OrderBy(x => x.Id).ToListAsync();
                    string finalFormattedMenu = FormatMenuItems(items);

                    string diffText = "```text\n(無法比對變動)\n```";

                    if (botConfig != null)
                    {
                        // 1. 取得舊版與新版的純文字 (利用正則表達式過濾掉 ANSI 色碼與 ``` 標籤)
                        string oldMenu = botConfig.RewardMenuContent ?? "";
                        var oldLines = oldMenu.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                                              .Where(l => !l.Contains("```"))
                                              .Select(l => System.Text.RegularExpressions.Regex.Replace(l, @"\x1B\[[^m]*m", "").Trim())
                                              .Where(l => !string.IsNullOrEmpty(l))
                                              .ToList();

                        var newLines = finalFormattedMenu.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                                              .Where(l => !l.Contains("```"))
                                              .Select(l => System.Text.RegularExpressions.Regex.Replace(l, @"\x1B\[[^m]*m", "").Trim())
                                              .Where(l => !string.IsNullOrEmpty(l))
                                              .ToList();

                        // 2. 進行差異比對 (抓出新增與移除的項目)
                        var added = newLines.Except(oldLines).ToList();
                        var removed = oldLines.Except(newLines).ToList();

                        if (added.Any() || removed.Any())
                        {
                            var sbDiff = new StringBuilder();
                            sbDiff.AppendLine("```diff");
                            foreach (var r in removed) sbDiff.AppendLine($"- {r}");
                            foreach (var a in added) sbDiff.AppendLine($"+ {a}");
                            sbDiff.AppendLine("```");
                            diffText = sbDiff.ToString();
                        }
                        else
                        {
                            diffText = "```text\n(無實質項目變動，僅重新發布排版)\n```";
                        }

                        // 3. 儲存新資料
                        botConfig.RewardMenuContent = finalFormattedMenu;
                        botConfig.RewardMenuUpdateTime = Program.GetTaipeiTime();
                        await db.SaveChangesAsync();
                    }

                    string timeStr = botConfig?.RewardMenuUpdateTime?.ToString("yyyy年MM月dd日HH:mm") ?? Program.GetTaipeiTime().ToString("yyyy年MM月dd日HH:mm");
                    await component.UpdateAsync(m => { m.Content = $"✅ **發布成功！** 線上 MENU 已更新。"; m.Components = null; });

                    // 原頻道內的公告維持原樣發布完整的彩色 MENU
                    string publicMsg = $"✅ **MENU 已成功更新發布！**\n> 執行者：{executor.Mention}\n\n{finalFormattedMenu}\n\n> *此為 {timeStr} 公告*";
                    await component.Channel.SendMessageAsync(publicMsg);

                    // 🚨 修改重點：這裡的 true 改為 false (不 Ping 管理員)，並且改用 Username 避免 Ping 到執行者，最後附上剛剛比對出來的 diffText 變動區塊
                    _ = BroadcastToAdminChannelsAsync(gid, $"📜 **【MENU 更新公告】**\n> 執行者：**{executor.Username}**\n> 此次 MENU 變動內容如下：\n{diffText}\n> *(大家可使用 `/menu` 查看完整最新品項)*", false);
                    return;
                }
                // 👇 新增：處理語音結算面板的下拉式選單與按鈕
                if (id.StartsWith("evadd_") || id.StartsWith("evrem_"))
                {
                    string sessId = id.Split('_')[1];
                    if (!_pendingDistributions.TryGetValue(sessId, out var pending)) { await component.RespondAsync("❌ 面板已過期。", ephemeral: true); return; }
                    if (executor.Id != pending.AdminId) { await component.RespondAsync("❌ 僅限發起此活動的長官操作。", ephemeral: true); return; }

                    var selectedIds = component.Data.Values.Select(ulong.Parse).ToList();
                    if (id.StartsWith("evadd_")) foreach (var uid in selectedIds) pending.FinalUserIds.Add(uid);
                    else foreach (var uid in selectedIds) pending.FinalUserIds.Remove(uid);

                    await UpdateEventDashboardAsync(component, sessId, executor.Guild);
                    return;
                }

                if (id.StartsWith("evcancel_"))
                {
                    string sessId = id.Split('_')[1];
                    _pendingDistributions.Remove(sessId);
                    await component.UpdateAsync(m => { m.Content = "❌ 此語音活動結算已作廢。"; m.Components = null; });
                    return;
                }

                if (id.StartsWith("evok_"))
                {
                    string sessId = id.Split('_')[1];
                    if (!_pendingDistributions.TryGetValue(sessId, out var pending)) { await component.RespondAsync("❌ 面板已失效。", ephemeral: true); return; }
                    if (executor.Id != pending.AdminId) { await component.RespondAsync("❌ 僅限發起此活動的長官操作。", ephemeral: true); return; }

                    await component.UpdateAsync(m => { m.Content = "⏳ 正在批次發放點數並推播，請稍候..."; m.Components = null; });

                    int successCount = 0;
                    List<string> failedNames = new();

                    // 開始跑批次發放
                    foreach (var uid in pending.FinalUserIds)
                    {
                        var targetUser = executor.Guild.GetUser(uid);
                        if (targetUser == null) continue;

                        string name = targetUser.Nickname ?? targetUser.Username;
                        if (name.Contains("]")) name = name.Substring(name.LastIndexOf(']') + 1).Trim();

                        var rec = await db.UserPoints.FirstOrDefaultAsync(u => u.GuildId == gid && u.RobloxUsername.ToLower() == name.ToLower());
                        if (rec == null)
                        {
                            // 嚴格檢查模式：確保此人有在 Roblox 群組內
                            if (!await VerifyUserInRobloxGroup(name, botConfig.RobloxGroupId)) { failedNames.Add(name); continue; }
                            rec = new UserPoint { GuildId = gid, RobloxUsername = name, Points = 0 };
                            db.UserPoints.Add(rec);
                        }

                        rec.Points += pending.Points;
                        db.PointLogs.Add(new PointLog { GuildId = gid, RobloxUsername = name, AdminName = executor.Username, PointsAdded = pending.Points, Reason = pending.Reason, Timestamp = Program.GetTaipeiTime() });
                        successCount++;
                    }

                    await db.SaveChangesAsync();
                    _ = UpdateGoogleSheetAsync(gid);
                    _pendingDistributions.Remove(sessId);

                    // 👇 計算用於最終推播的總時長 (精簡到分鐘即可)
                    var duration = pending.EndTime - pending.StartTime;
                    string durationStr = $"{(int)duration.TotalHours} 小時 {duration.Minutes} 分鐘";

                    // 👇 在廣播訊息中加入「活動總時長」
                    string resultMsg = $"✅ **語音活動點數發放完成！**\n>  **成功發放對象：** {successCount} 人\n>  **每人獲得：** {pending.Points} 點\n>  **原因：** {pending.Reason}\n>  **活動總時長：** {durationStr}\n>  **登記長官：** {executor.Mention}";
                    if (failedNames.Any()) resultMsg += $"\n\n⚠️ **下列人員發放失敗 (未綁定Roblox或不在群組)：**\n`{string.Join(", ", failedNames)}`";

                    // 發布公開廣播 (不限制僅長官可見)
                    await component.Channel.SendMessageAsync(resultMsg);

                    // 推播給總部與 Log 頻道的訊息也加上時長
                    bool isAnomaly = pending.Points >= 100;
                    _ = BroadcastToAdminChannelsAsync(gid, $"負責人 **{executor.Username}** 已完成語音活動批次發放\n>  活動：**{pending.Reason}**\n>  總時長：**{durationStr}**\n>  參與人數：**{successCount}** 人\n>  每人獲得：**{pending.Points}** 點", isAnomaly);
                    return;
                }
                // ================================================================
                // 處理強制同步試算表的按鈕
                if (id.StartsWith("force_sync_sheet_"))
                {
                    if (!isAdmin)
                    {
                        await component.RespondAsync("❌ 權限不足，僅限管理員執行。", ephemeral: true);
                        return;
                    }

                    // 告訴使用者正在處理中 (ephemeral 代表只有點擊的人看得到)
                    await component.RespondAsync("⏳ 正在將最新資料上傳至 Google 試算表，請稍候...", ephemeral: true);

                    // 執行同步
                    await UpdateGoogleSheetAsync(gid);

                    // 更新狀態
                    await component.FollowupAsync("✅ **同步完成！** 您現在可以點擊上方的連結前往查看最新數據。", ephemeral: true);
                    return;
                }
            }
        }
        // 💡 回傳型態改為 (int added, int removed)
        private async Task<(int added, int removed)> SyncGroupMembersAsync(BotDbContext db, ulong guildId, string groupId)
        {
            int maxRank = 255;
            switch (groupId)
            {
                case "13549943": maxRank = 239; break;
                case "13662982": maxRank = 120; break;
                case "16223475": maxRank = 55; break;
            }
            int minRank = 1;

            string cursor = "";
            bool hasMore = true;
            int addedCount = 0;
            int removedCount = 0;

            // 改用 Dictionary 以便快速讀取與更新現有使用者的職位
            var existingDbUsers = await db.UserPoints.Where(u => u.GuildId == guildId).ToListAsync();
            var existingDict = existingDbUsers.ToDictionary(u => u.RobloxUsername.ToLower());
            var validActiveUsers = new HashSet<string>();

            while (hasMore)
            {
                string url = $"https://groups.roblox.com/v1/groups/{groupId}/users?limit=100";
                if (!string.IsNullOrEmpty(cursor)) url += $"&cursor={cursor}";

                var res = await _http.GetAsync(url);
                if (!res.IsSuccessStatusCode)
                {
                    Console.WriteLine("⚠️ Roblox API 請求失敗或已達限制，中斷同步以保護資料庫。");
                    return (addedCount, removedCount);
                }

                var json = await res.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                var data = root.GetProperty("data");
                foreach (var item in data.EnumerateArray())
                {
                    string username = item.GetProperty("user").GetProperty("username").GetString();
                    var roleProp = item.GetProperty("role"); // 取得 role 物件
                    int rank = roleProp.GetProperty("rank").GetInt32();
                    string roleName = roleProp.GetProperty("name").GetString(); // 取得原始職位名稱

                    if (rank <= maxRank && rank >= minRank)
                    {
                        validActiveUsers.Add(username.ToLower());

                        // 🔍 自動擷取 [ ] 內的文字 (例如：[大隊長-勤務支援中隊])
                        string parsedRank = roleName;
                        if (!string.IsNullOrEmpty(roleName) && roleName.Contains("[") && roleName.Contains("]"))
                        {
                            int start = roleName.IndexOf('[');
                            int end = roleName.IndexOf(']', start);
                            if (start != -1 && end != -1)
                            {
                                parsedRank = roleName.Substring(start, end - start + 1);
                            }
                        }

                        // 如果玩家已經在資料庫，檢查並更新他的職位
                        if (existingDict.TryGetValue(username.ToLower(), out var existingUser))
                        {
                            if (existingUser.RobloxRank != parsedRank) existingUser.RobloxRank = parsedRank;
                        }
                        else
                        {
                            // 如果不在資料庫，就新增並填入職位
                            db.UserPoints.Add(new UserPoint { GuildId = guildId, RobloxUsername = username, Points = 0, RobloxRank = parsedRank });
                            existingDict[username.ToLower()] = new UserPoint { RobloxRank = parsedRank }; // 避免同頁面重複新增
                            addedCount++;
                        }
                    }
                }

                // 👇 將原本判斷下一頁的程式碼，替換成下面這段：
                if (root.TryGetProperty("nextPageCursor", out var nextCursor) && nextCursor.ValueKind == JsonValueKind.String)
                {
                    cursor = nextCursor.GetString();
                    // 🛡️ 新增這行：強制等待 0.5 秒 (500 毫秒) 再去抓下一頁
                    await Task.Delay(500);
                }
                else
                {
                    hasMore = false;
                }
            } // <-- 這是 while 迴圈的右括號

            var usersToDelete = await db.UserPoints.Where(u => u.GuildId == guildId).ToListAsync();
            var finalDeleteList = usersToDelete.Where(u => !validActiveUsers.Contains(u.RobloxUsername.ToLower())).ToList();

            if (finalDeleteList.Any())
            {
                db.UserPoints.RemoveRange(finalDeleteList);
                removedCount = finalDeleteList.Count;
                Console.WriteLine($"🗑️ 已自動清理 {removedCount} 位不符合權重或退群的成員。");
            }

            // 👇 拔除條件判斷，讓職位更新可以順利存進資料庫！
            // (EF Core 很聰明，如果資料沒變它就不會寫入，所以直接呼叫是絕對安全的)
            await db.SaveChangesAsync();

            return (addedCount, removedCount);

          
        }

        private async Task<bool> VerifyUserInRobloxGroup(string username, string groupId)
        {
            try
            {
                var userReq = new { usernames = new[] { username }, excludeBannedUsers = true };
                var content = new StringContent(JsonSerializer.Serialize(userReq), Encoding.UTF8, "application/json");

                // 【已修復】確保這裡只有純文字的 https 開頭，絕對沒有中括號 [ ]
                var userRes = await _http.PostAsync("https://users.roblox.com/v1/usernames/users", content);
                var userJson = await userRes.Content.ReadAsStringAsync();
                var data = JsonDocument.Parse(userJson).RootElement.GetProperty("data");
                if (data.GetArrayLength() == 0) return false;
                long userId = data[0].GetProperty("id").GetInt64();

                // 【已修復】確保這裡只有純文字的 https 開頭，絕對沒有中括號 [ ]
                var groupRes = await _http.GetAsync($"https://groups.roblox.com/v1/users/{userId}/groups/roles");
                var groupJson = await groupRes.Content.ReadAsStringAsync();
                return groupJson.Contains($"\"id\":{groupId}");
            }
            catch { return false; }
        }
        // 自動將管理員的純文字轉換為精美排版的 MENU
        // 自動將管理員的純文字轉換為精美排版與 ANSI 色塊的 MENU


        // 計算字串在 Discord 顯示的真實寬度 (中文字算寬度 2)
        private int GetDisplayWidth(string s)
        {
            int length = 0;
            foreach (char c in s) length += c > 0xFF ? 2 : 1;
            return length;
        }

        // 自動計算該伺服器擁有「機器人管理員身分組」的總人數
        private int GetAdminCount(SocketGuild guild, BotConfig config)
        {
            if (config == null) return 0;
            var roleIds = new List<ulong>();
            if (!string.IsNullOrEmpty(config.AdminRoleIds))
            {
                roleIds = config.AdminRoleIds.Split(',').Where(s => !string.IsNullOrEmpty(s)).Select(ulong.Parse).ToList();
            }
            if (!roleIds.Any() && config.AdminRoleId != 0) roleIds.Add(config.AdminRoleId);

            int count = 0;
            foreach (var user in guild.Users)
            {
                if (user.Roles.Any(r => roleIds.Contains(r.Id))) count++;
            }
            return count;
        }
        // 補齊空格，讓所有中英文都能完美對齊
        private string PadRightCustom(string s, int totalWidth)
        {
            int currentWidth = GetDisplayWidth(s);
            if (currentWidth >= totalWidth) return s;
            return s + new string(' ', totalWidth - currentWidth);
        }
        // 統一驗證使用者是否具備管理員權限
        private bool IsAdmin(SocketGuildUser user, BotConfig config)
        {
            if (user.GuildPermissions.Administrator) return true; // 伺服器管理員直接放行
            if (config == null) return false;

            // 檢查是否有多重身分組設定
            if (!string.IsNullOrEmpty(config.AdminRoleIds))
            {
                var allowedIds = config.AdminRoleIds.Split(',').Select(s => ulong.TryParse(s, out var id) ? id : 0).Where(id => id != 0);
                return user.Roles.Any(r => allowedIds.Contains(r.Id));
            }
            // 兼容舊版的單一身分組檢查
            return user.Roles.Any(r => r.Id == config.AdminRoleId);
        }
        // 👇 新增這段：負責計算進出入時間與紀錄文字日誌
        private Task HandleVoiceStateUpdatedAsync(SocketUser user, SocketVoiceState oldState, SocketVoiceState newState)
        {
            // 情況 A：離開了原本的語音頻道
            if (oldState.VoiceChannel != null && _activeEvents.TryGetValue(oldState.VoiceChannel.Id, out var leftEvt))
            {
                if (leftEvt.Participants.TryGetValue(user.Id, out var p) && p.LastJoinTime.HasValue)
                {
                    p.TotalSeconds += (Program.GetTaipeiTime() - p.LastJoinTime.Value).TotalSeconds;
                    p.LastJoinTime = null; // 標記為不在頻道內
                    leftEvt.ActionLogs.Add($"[{Program.GetTaipeiTime():HH:mm:ss}] 🔴 {user.Username} 離開了語音");
                }
            }
            // 情況 B：加入了正在追蹤的頻道
            if (newState.VoiceChannel != null && _activeEvents.TryGetValue(newState.VoiceChannel.Id, out var joinEvt))
            {
                if (!joinEvt.Participants.ContainsKey(user.Id))
                    joinEvt.Participants[user.Id] = new VoiceParticipant();

                joinEvt.Participants[user.Id].LastJoinTime = Program.GetTaipeiTime();
                joinEvt.ActionLogs.Add($"[{Program.GetTaipeiTime():HH:mm:ss}] 🟢 {user.Username} 加入了語音");
            }
            return Task.CompletedTask;
        }
        private async Task BroadcastToAdminChannelsAsync(ulong sourceGuildId, string message, bool isAnomaly = false)
        {
            using var db = new BotDbContext(_configuration);
            var sourceConfig = await db.Configs.FindAsync(sourceGuildId);
            if (sourceConfig == null || string.IsNullOrEmpty(sourceConfig.ServerCode)) return;
            // 👇 ================= 新增這段 ================= 👇
            // 1. 推播給【本伺服器】指定的 Log 頻道
            if (sourceConfig.MasterLogChannelId.HasValue && sourceConfig.MasterLogChannelId.Value != 0)
            {
                if (_client.GetChannel(sourceConfig.MasterLogChannelId.Value) is IMessageChannel localChannel)
                {
                    try
                    {
                        // 如果是異常大額變動，就用紅燈警告並 Ping 管理員 1 (AdminRoleId)
                        if (isAnomaly)
                        {
                            await localChannel.SendMessageAsync($"🚨 **【點數變動異常通知】** <@&{sourceConfig.AdminRoleId}>\n> {message}");
                        }
                        else
                        {
                            await localChannel.SendMessageAsync($"📜 **[點數變動通知]**\n> {message}");
                        }
                    }
                    catch { /* 如果機器人沒有該頻道的發言權限，則安全忽略 */ }
                }
            }
            // 👆 ========================================== 👆
            // 👇 新增這段：透過群組 ID 判斷單位名稱
            string unitName = sourceConfig.RobloxGroupId switch
            {
                "13549943" => "憲兵",
                "13662982" => "裝甲",
                "16223475" => "航特",
                _ => "未知"
            };

            var adminConfigs = await db.AdminChannels.ToListAsync();
            foreach (var ac in adminConfigs)
            {
                // 如果這個 Admin 頻道有監控發生變動的伺服器編號
                if (ac.TargetServerCodes.Contains(sourceConfig.ServerCode))
                {
                    if (_client.GetChannel(ac.ChannelId) is IMessageChannel channel)
                    {
                        try
                        {
                            // 👇 修改這行：讓推播訊息顯示為 "裝甲單位點數通知[資料庫3A3F3]" 的格式
                            await channel.SendMessageAsync($" **{unitName}單位點數通知[資料庫{sourceConfig.ServerCode}]**\n> {message}");
                        }
                        catch { /* 若沒有頻道發言權限則忽略 */ }
                    }
                }
            }
        }
        private async Task UpdateGoogleSheetAsync(ulong guildId)
        {
            try
            {
                using var db = new BotDbContext(_configuration);
                var config = await db.Configs.FindAsync(guildId);

                // 如果沒有設定群組ID，則無法判斷單位，直接中斷
                if (config == null || string.IsNullOrEmpty(config.RobloxGroupId)) return;

                string individualSheetId = "";
                string tabName = "";
                // 這是你提供的「三單位總覽試算表 ID」
                string combinedSheetId = "1v8TmW04kJwUKPFeghK5OkNzyj65XSrcHXHWGfsokTQ8";

                // 根據綁定的 Roblox 群組 ID，自動對應專屬試算表與分頁名稱
                switch (config.RobloxGroupId)
                {
                    case "13549943": // 憲兵
                        individualSheetId = "1cSaJJXq75MJ479RV5KcZ8oYBHxdbexC5f0XQtzd7Bwo";
                        tabName = "憲兵";
                        break;
                    case "13662982": // 裝甲
                        individualSheetId = "1k08y_tUTA0_eS_W9zkGQzHU3OSLHcVlzKtTXvjzUxB4";
                        tabName = "裝甲";
                        break;
                    case "16223475": // 航特
                        individualSheetId = "1rOXjrRdBD1S77guAByBaACgxToReHMUhSiU2WE1_zrE";
                        tabName = "航特";
                        break;
                    default:
                        // 若未來有其他單位，且有使用 /bind-sheet 綁定自訂 ID 時的備用方案
                        if (!string.IsNullOrEmpty(config.GoogleSheetId))
                        {
                            individualSheetId = config.GoogleSheetId;
                            tabName = ""; // 無對應總覽分頁
                        }
                        else return;
                        break;
                }

                // 從資料庫抓取最新的排行榜資料
                var users = await db.UserPoints.Where(u => u.GuildId == guildId).OrderByDescending(u => u.Points).ToListAsync();

                // 準備要寫入 Google 試算表的資料結構 (3個欄位)
                var valueRange = new ValueRange();
                var oblist = new List<IList<object>>()
                {
                    new List<object>() { "職位", "人員名稱", "目前點數", "最後變動時間" } // 👈 加入職位表頭
                };

                string updateTime = Program.GetTaipeiTime().ToString("yyyy-MM-dd HH:mm:ss");
                foreach (var u in users)
                {
                    // 寫入四欄資料：職位 (若空則顯示未知)、名稱、點數、時間
                    oblist.Add(new List<object>() { u.RobloxRank ?? "尚未同步", u.RobloxUsername, u.Points, updateTime });
                }
                valueRange.Values = oblist;

                // 讀取 Google 服務帳戶金鑰
                GoogleCredential credential;
                using (var stream = new System.IO.FileStream("google-credentials.json", System.IO.FileMode.Open, System.IO.FileAccess.Read))
                {
                    credential = GoogleCredential.FromStream(stream).CreateScoped(SheetsService.Scope.Spreadsheets);
                }

                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "ROCA Point Bot"
                });

                // 🟢 動作 1：更新該單位的「專屬試算表」
                if (!string.IsNullOrEmpty(individualSheetId))
                {
                    // 改為清空 A 到 D 欄
                    await service.Spreadsheets.Values.Clear(new ClearValuesRequest(), individualSheetId, "A:D").ExecuteAsync();
                    var updateRequest = service.Spreadsheets.Values.Update(valueRange, individualSheetId, "A1");
                    updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    await updateRequest.ExecuteAsync();
                }

                // 🟢 動作 2：同步更新「三單位總覽試算表」的對應分頁
                if (!string.IsNullOrEmpty(tabName))
                {
                    // 指定範圍改為 A:D
                    string range = $"{tabName}!A:D";
                    await service.Spreadsheets.Values.Clear(new ClearValuesRequest(), combinedSheetId, range).ExecuteAsync();

                    var updateCombinedReq = service.Spreadsheets.Values.Update(valueRange, combinedSheetId, $"{tabName}!A1");
                    updateCombinedReq.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                    await updateCombinedReq.ExecuteAsync();
                }

                Console.WriteLine($"✅ 伺服器 {guildId} (單位:{tabName}) 的專屬試算表與總覽試算表同步成功！");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Google Sheet 同步失敗: {ex.Message}");
            }
        }
    }

    public class BotDbContext : DbContext
    {
        private readonly IConfiguration _config;
        public BotDbContext(IConfiguration config) { _config = config; }

        public DbSet<UserPoint> UserPoints { get; set; }
        // 2. 在 BotDbContext 類別中，加入這行 DbSet (加在 UserPoints 附近)
        public DbSet<RewardMenuItem> RewardMenuItems { get; set; }
        public DbSet<PointLog> PointLogs { get; set; }
        public DbSet<BotConfig> Configs { get; set; }
        public DbSet<AdminChannelConfig> AdminChannels { get; set; }
     
        protected override void OnConfiguring(DbContextOptionsBuilder options)
        {
            string conn = _config["DbConnection"];
            if (!string.IsNullOrEmpty(conn)) options.UseSqlServer(conn);
            else options.UseSqlite("Data Source=rocapoints.db");
        }
    }

    public class UserPoint { [Key] public int Id { get; set; } public ulong GuildId { get; set; } public string RobloxUsername { get; set; } public int Points { get; set; } public string? RobloxRank { get; set; } }

    public class PointLog { [Key] public int Id { get; set; } public ulong GuildId { get; set; } public string RobloxUsername { get; set; } public string AdminName { get; set; } public int PointsAdded { get; set; } public string Reason { get; set; } public DateTime Timestamp { get; set; } public bool IsDeleted { get; set; } }
    // 👇 新增這行：加入 Admin 頻道設定表
      

    [Table("GuildConfigs")]
    public class BotConfig
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public ulong GuildId { get; set; }
        public string RobloxGroupId { get; set; }
        public ulong AdminRoleId { get; set; }

        // 👇 新增的兩個欄位：伺服器編號與總部頻道 ID
        public string? ServerCode { get; set; }
        // 👇 新增這個欄位：用來儲存本伺服器專屬的 Log 頻道 ID
        public ulong? MasterLogChannelId { get; set; }
        // 👇 新增這個欄位：用來儲存多個身分組 ID (以逗號分隔)
        public string? AdminRoleIds { get; set; }
        // 👇 新增這個欄位：用來儲存各單位的獨立 Google Sheet ID
        public string? GoogleSheetId { get; set; }
        // 在 BotConfig 類別中加入這兩行：
        public string? RewardMenuContent { get; set; }
        public DateTime? RewardMenuUpdateTime { get; set; }
    }
    // 👇 新增這個類別：記錄哪個頻道負責監聽哪些伺服器編號
    public class AdminChannelConfig
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public ulong ChannelId { get; set; }
        public ulong GuildId { get; set; }
        public string TargetServerCodes { get; set; } // 存放要監控的編號，例如 "A1B2C,X9Y8Z"

        // 👇 新增這行：儲存國防部專屬查詢身分組 ID
        public ulong? MndAdminRoleId { get; set; }
        public ulong? MndAdminRoleId2 { get; set; } // 👈 新增這行：儲存第二個長官身分組
    }
    // 1. 在檔案最底下加入這個新的類別
    public class RewardMenuItem
    {
        [Key] public int Id { get; set; }
        public ulong GuildId { get; set; }
        public string ItemName { get; set; }
        public int Points { get; set; }
        public string? Note { get; set; }
    }
    // 👇 新增：語音活動追蹤專用的類別
    public class VoiceEventSession
    {
        public ulong GuildId { get; set; }
        public ulong ChannelId { get; set; }
        public ulong AdminId { get; set; }
        public string Reason { get; set; }
        public DateTime StartTime { get; set; }
        public Dictionary<ulong, VoiceParticipant> Participants { get; set; } = new();
        public List<string> ActionLogs { get; set; } = new();
    }
    public class VoiceParticipant
    {
        public DateTime? LastJoinTime { get; set; }
        public double TotalSeconds { get; set; }
    }
    public class PendingEventDistribution
    {
        public ulong GuildId { get; set; }
        public int Points { get; set; }
        public string Reason { get; set; }
        public ulong AdminId { get; set; }
        public HashSet<ulong> FinalUserIds { get; set; } = new();
        public List<string> ActionLogs { get; set; } = new();
        // 👇 新增這兩個屬性來記錄活動總時間
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
    }
}