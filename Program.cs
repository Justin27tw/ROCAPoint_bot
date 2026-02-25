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

        public DiscordBotService(IConfiguration config)
        {
            _configuration = config;
            _discordToken = config["DiscordToken"];
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            var config = new DiscordSocketConfig { GatewayIntents = GatewayIntents.AllUnprivileged | GatewayIntents.GuildMembers };
            _client = new DiscordSocketClient(config);
            _client.Ready += RegisterCommandsAsync;
            _client.SlashCommandExecuted += HandleSlashCommandAsync;
            _client.InteractionCreated += HandleInteractionAsync;

            await _client.LoginAsync(TokenType.Bot, _discordToken);
            await _client.StartAsync();

            _ = Task.Run(async () =>
            {
                try
                {
                    using var db = new BotDbContext(_configuration);
                    await db.Database.EnsureCreatedAsync();

                    // 💡 自動嘗試在現有資料表中加入新欄位 (用 Try-Catch 防止重複加入報錯)
                    try { await db.Database.ExecuteSqlRawAsync("ALTER TABLE GuildConfigs ADD ServerCode NVARCHAR(50) NULL"); } catch { }
                    try { await db.Database.ExecuteSqlRawAsync("ALTER TABLE GuildConfigs ADD MasterLogChannelId BIGINT NULL"); } catch { }

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
                .AddOption("admin_role", ApplicationCommandOptionType.Role, "管理身分組", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("sync-members").WithDescription("🔄 手動同步 Roblox 群組的最新成員至資料庫").Build(),
                new SlashCommandBuilder().WithName("points").WithDescription("📊 查詢點數").AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("addpoint").WithDescription("➕ 發放點數").AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true).AddOption("points", ApplicationCommandOptionType.Integer, "點數數量", isRequired: true).AddOption("reason", ApplicationCommandOptionType.String, "原因備註", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("history").WithDescription("📜 查詢玩家近期紀錄").AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("viewall").WithDescription("🏆 點數總排行榜 (顯示所有成員)").Build(),
                new SlashCommandBuilder().WithName("del-record").WithDescription("🚨 刪除單筆紀錄").AddOption("id", ApplicationCommandOptionType.Integer, "紀錄 ID", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("unbind-roca").WithDescription("🔓 解除綁定設定").Build(),
                new SlashCommandBuilder().WithName("clear-all-data").WithDescription("🔥 【極度危險】刪除本伺服器所有資料").Build(),
                new SlashCommandBuilder().WithName("daily-records").WithDescription("📅 查詢指定日期的所有點數紀錄").AddOption("date", ApplicationCommandOptionType.String, "格式: YYYY-MM-DD", isRequired: true).Build(),
                // 👇 新增這行：註冊 status 指令
                new SlashCommandBuilder().WithName("status").WithDescription("🟢 檢查機器人目前連線與運作狀態").Build(),
                // 👇 新增這兩個指令
                new SlashCommandBuilder().WithName("admin-setup").WithDescription("🔒 國防部專用：將此頻道設為 Admin 頻道以監控其他單位").AddOption("server_codes", ApplicationCommandOptionType.String, "要監控的伺服器編號(多個用逗號分隔，如 A1B2,C3D4)", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("admin-view").WithDescription("👁️ 國防部專用：查看特定單位的總排行榜").AddOption("server_code", ApplicationCommandOptionType.String, "目標伺服器編號", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("my-code").WithDescription("🔑 查詢本伺服器的專屬資料庫編號 (交給國防部綁定通知用)").Build(),
            };
            try { await _client.BulkOverwriteGlobalApplicationCommandsAsync(commands.ToArray()); } catch (Exception ex) { Console.WriteLine(ex.Message); }
        }

        private async Task HandleSlashCommandAsync(SocketSlashCommand command)
        {
            if (!(command.Channel is SocketGuildChannel guildChannel)) return;
            ulong gid = guildChannel.Guild.Id;

            try
            {
                bool isEphemeral = command.Data.Name == "setup-roca" || command.Data.Name == "unbind-roca" || command.Data.Name == "clear-all-data" || command.Data.Name == "sync-members" || command.Data.Name == "points" || command.Data.Name == "history" || command.Data.Name == "my-code";
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
                            var rRole = (SocketRole)command.Data.Options.First(x => x.Name == "admin_role").Value;

                            string newCode = Guid.NewGuid().ToString("N").Substring(0, 5).ToUpper();

                            if (botConfig == null)
                            {
                                botConfig = new BotConfig { GuildId = gid, RobloxGroupId = rId, AdminRoleId = rRole.Id, ServerCode = newCode };
                                db.Configs.Add(botConfig);
                            }
                            else
                            {
                                botConfig.RobloxGroupId = rId; botConfig.AdminRoleId = rRole.Id;
                                if (string.IsNullOrEmpty(botConfig.ServerCode)) botConfig.ServerCode = newCode;
                            }
                            await db.SaveChangesAsync();

                            await command.FollowupAsync($"✅ 已綁定 Roblox 群組 `{rId}`。\n🔑 **本伺服器專屬資料庫編號：`{botConfig.ServerCode}`** (請妥善保存，交給國防部設定 Admin 頻道使用)。\n正在背景抓取群組成員，請稍候...");
                            var syncResult = await SyncGroupMembersAsync(db, gid, rId);
                            await command.Channel.SendMessageAsync($"🔄 群組名單同步完成！\n> 🟢 新增了 **{syncResult.added}** 位新成員 (預設為 0 點)\n> 🔴 移除了 **{syncResult.removed}** 位已退群或不符資格成員。");
                            break;
                        }

                    case "sync-members":
                        {
                            if (botConfig == null) { await command.FollowupAsync("❌ 請先使用 `/setup-roca` 綁定群組。"); return; }
                            if (!((SocketGuildUser)command.User).Roles.Any(r => r.Id == botConfig.AdminRoleId)) { await command.FollowupAsync("❌ 權限不足。"); return; }

                            await command.FollowupAsync("⏳ 正在與 Roblox 同步群組成員名單...");
                            var syncResult = await SyncGroupMembersAsync(db, gid, botConfig.RobloxGroupId);
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
                            bool isCodeAdmin = myCodeExec.Roles.Any(r => r.Id == botConfig.AdminRoleId) || myCodeExec.GuildPermissions.Administrator;

                            if (!isCodeAdmin)
                            {
                                await command.FollowupAsync("❌ 權限不足！只有管理員可以查詢伺服器專屬編號。");
                                return;
                            }

                            await command.FollowupAsync($"🔑 **本伺服器專屬資料庫編號為：** `{botConfig.ServerCode}`\n> 請將此編號交給國防部，以便國防部進行跨伺服器監控設定。");
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

                            bool isFirst = true;
                            foreach (var chunk in chunks)
                            {
                                if (isFirst) { await command.FollowupAsync(chunk); isFirst = false; }
                                else { await command.Channel.SendMessageAsync(chunk); }
                            }
                            break;
                        }

                    case "addpoint":
                        {
                            if (botConfig == null) { await command.FollowupAsync("❌ 未設定。請先使用 /setup-roca"); break; }
                            var exec = (SocketGuildUser)command.User;
                            if (!exec.Roles.Any(r => r.Id == botConfig.AdminRoleId) && exec.Id != guildChannel.Guild.OwnerId) { await command.FollowupAsync("❌ 權限不足。"); return; }

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

                            // 【修改後的訊息格式：移除反引號 ` ，改用粗體 ** 】
                            string addMsg = $">  [點數發放] 負責人：**{command.User.Username}**\n" +
                                            $">  成功發放 **{pts}** 點給 **{name}**\n" +
                                            $">  目前總計：**{rec.Points}** 點\n" +
                                            $">  備註：{reason} \n" +
                                            $">  紀錄編號：{newLog.Id}";  // 這裡也修正為 {newLog.Id}
                            await command.FollowupAsync(addMsg);

                            _ = BroadcastToAdminChannelsAsync(gid, $"負責人 **{command.User.Username}** 已發放點數\n" +
                                        $">  被登記人：**{name}** 獲得 **{pts}** 點\n" +
                                        $">  目前總計：**{rec.Points}** 點\n" +
                                        $">  備註：{reason}\n" +
                                        $">  紀錄編號：**{newLog.Id}** (若需撤銷請使用 /del-record)");
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
                                string ptsRaw = $"+{l.PointsAdded}".PadLeft(5);
                                string adminStr = l.AdminName.PadRight(12);

                                string idStr = $"\u001b[34m{idRaw}\u001b[0m";
                                string ptsStr = $"\u001b[32m{ptsRaw}\u001b[0m";

                                sb.AppendLine($"{idStr} \u001b[30m{timeStr}\u001b[0m | ➔ {ptsStr} 點 | 登記: {adminStr} | 原因: {l.Reason}");
                            }
                            sb.AppendLine("```");
                            await command.FollowupAsync(sb.ToString());
                            break;
                        }

                    case "del-record":
                        {
                            if (botConfig == null) { await command.FollowupAsync("⚠️ 請先設定。"); return; }
                            if (!((SocketGuildUser)command.User).Roles.Any(r => r.Id == botConfig.AdminRoleId)) { await command.FollowupAsync("❌ 權限不足。"); return; }

                            int logId = Convert.ToInt32((long)command.Data.Options.First().Value);
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
                            targetLog.Reason += $" (已由 {command.User.Username} 撤銷)";

                            await db.SaveChangesAsync();

                            string delMsg = $" **[紀錄撤銷]** 負責人：**{command.User.Username}**\n" +
                                            $"> 成功撤銷了紀錄 #{logId}\n" +
                                            $">  被扣點人：{targetLog.RobloxUsername}\n" +
                                            $">  已扣除：{targetLog.PointsAdded} 點\n" +
                                            $">  **目前剩餘：{currentPoints} 點**";

                            await command.FollowupAsync(delMsg);
                            _ = BroadcastToAdminChannelsAsync(gid, $"負責人 **{command.User.Username}** 撤銷了紀錄 `#{logId}`，扣除 {targetLog.PointsAdded} 點。\n>  人員：{targetLog.RobloxUsername}\n>  剩餘：{currentPoints}點");
                            break;
                        }

                    case "daily-records":
                        {
                            string dateStr = (string)command.Data.Options.First().Value;
                            if (!DateTime.TryParse(dateStr, out DateTime dt)) { await command.FollowupAsync("❌ 日期格式錯誤。"); break; }
                            var dLogs = await db.PointLogs.Where(l => l.GuildId == gid && l.Timestamp.Date == dt.Date && !l.IsDeleted).ToListAsync();
                            if (!dLogs.Any()) { await command.FollowupAsync("📭 該日無紀錄。"); break; }

                            var dSb = new StringBuilder($"#  {dt:yyyy-MM-dd} 發放紀錄清單\n```text\n");
                            foreach (var l in dLogs)
                            {
                                string idStr = $"[ID: {l.Id}]".PadRight(9);
                                string timeStr = l.Timestamp.ToString("HH:mm");
                                string nameStr = l.RobloxUsername.PadRight(18);
                                string ptsStr = $"+{l.PointsAdded}".PadLeft(5);
                                string adminStr = l.AdminName.PadRight(12);

                                dSb.AppendLine($"{idStr} {timeStr} | {nameStr} ➔ {ptsStr} 點 | 登記: {adminStr} | 原因: {l.Reason}");
                            }
                            dSb.AppendLine("```");
                            await command.FollowupAsync(dSb.ToString());
                            break;
                        }

                    case "clear-all-data":
                        {
                            if (!((SocketGuildUser)command.User).GuildPermissions.Administrator) { await command.FollowupAsync("❌ 限管理員。"); return; }
                            var btns = new ComponentBuilder().WithButton("確認刪除 (1/2)", $"clear_step1_{gid}", ButtonStyle.Danger).WithButton("取消", $"clear_cancel_{gid}", ButtonStyle.Secondary);
                            await command.FollowupAsync("⚠️ **【第一層確認】** 確定要清空所有資料嗎？", components: btns.Build());
                            break;
                        }

                    case "unbind-roca":
                        {
                            if (!((SocketGuildUser)command.User).GuildPermissions.Administrator) { await command.FollowupAsync("❌ 限管理員執行。"); return; }
                            if (botConfig != null) { db.Configs.Remove(botConfig); await db.SaveChangesAsync(); await command.FollowupAsync("🔓 已成功解除本伺服器的設定。"); }
                            else { await command.FollowupAsync("⚠️ 本伺服器尚未進行綁定。"); }
                            break;
                        }

                    case "points":
                        {
                            var targetUser = (SocketGuildUser)command.Data.Options.First(x => x.Name == "user").Value;
                            string targetName = targetUser.Nickname ?? targetUser.Username;
                            if (targetName.Contains("]")) targetName = targetName.Substring(targetName.LastIndexOf(']') + 1).Trim();
                            var userPoint = await db.UserPoints.FirstOrDefaultAsync(u => u.GuildId == gid && u.RobloxUsername.ToLower() == targetName.ToLower());
                            int currentPoints = userPoint != null ? userPoint.Points : 0;
                            await command.FollowupAsync($"📊 **{targetName}** 目前擁有 **{currentPoints}** 點。");
                            break;
                        }

                    case "admin-setup":
                        {
                            if (botConfig == null) { await command.FollowupAsync("❌ 請先在伺服器使用 `/setup-roca` 完成基本設定。"); return; }
                            if (!((SocketGuildUser)command.User).Roles.Any(r => r.Id == botConfig.AdminRoleId)) { await command.FollowupAsync("❌ 權限不足！只有設定好的 Admin 身分組可以設定監控頻道。"); return; }

                            string codes = ((string)command.Data.Options.First().Value).ToUpper();
                            var adminChannel = await db.AdminChannels.FindAsync(command.Channel.Id);

                            if (adminChannel == null) db.AdminChannels.Add(new AdminChannelConfig { ChannelId = command.Channel.Id, GuildId = gid, TargetServerCodes = codes });
                            else adminChannel.TargetServerCodes = codes;

                            await db.SaveChangesAsync();
                            await command.FollowupAsync($"✅ **已將此頻道設為 Admin 監控頻道！**\n📡 目前授權監控的單位編號：`{codes}`\n(只要這些單位資料庫有變動，就會自動回傳至此頻道)");
                            break;
                        }

                    case "admin-view":
                        {
                            var ac = await db.AdminChannels.FindAsync(command.Channel.Id);
                            if (ac == null) { await command.FollowupAsync("❌ 此頻道尚未被設定為 Admin 監控頻道。"); return; }

                            string targetCode = ((string)command.Data.Options.First().Value).ToUpper();
                            if (!ac.TargetServerCodes.Contains(targetCode)) { await command.FollowupAsync($"❌ 此頻道未被授權查看編號 `{targetCode}` 的資料庫。"); return; }

                            var targetConfig = await db.Configs.FirstOrDefaultAsync(c => c.ServerCode == targetCode);
                            if (targetConfig == null) { await command.FollowupAsync("❌ 找不到該編號的伺服器，請確認單位是否已綁定機器人。"); return; }

                            var targetData = await db.UserPoints.Where(u => u.GuildId == targetConfig.GuildId).OrderByDescending(u => u.Points).ToListAsync();
                            if (!targetData.Any()) { await command.FollowupAsync($"📭 目標編號 `{targetCode}` 尚無成員資料。"); break; }

                            var targetSb = new StringBuilder($"## 👁️ 國防部查閱：單位 [{targetCode}] 點數總覽\n```ansi\n");
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
                if (command.HasResponded) await command.FollowupAsync(errMsg, ephemeral: true);
                else await command.RespondAsync(errMsg, ephemeral: true);
            }
        }

        

        private async Task HandleInteractionAsync(SocketInteraction interaction)
        {
            if (interaction is SocketMessageComponent component)
            {
                string id = component.Data.CustomId;
                ulong gid = (ulong)component.GuildId;
                if (id.StartsWith("clear_cancel_")) { await component.UpdateAsync(m => { m.Content = "❌ 已取消。"; m.Components = null; }); return; }
                if (id.StartsWith("clear_step1_"))
                {
                    var btn2 = new ComponentBuilder().WithButton("⚠️ 最終確認 (2/2)", $"clear_step2_{gid}", ButtonStyle.Danger).WithButton("取消", $"clear_cancel_{gid}", ButtonStyle.Secondary);
                    await component.UpdateAsync(m => { m.Content = "🚨 **【最終警告】** 資料將永久刪除！"; m.Components = btn2.Build(); });
                    return;
                }
                if (id.StartsWith("clear_step2_"))
                {
                    using var db = new BotDbContext(_configuration);
                    db.UserPoints.RemoveRange(db.UserPoints.Where(u => u.GuildId == gid));
                    foreach (var l in db.PointLogs.Where(l => l.GuildId == gid)) l.IsDeleted = true;
                    await db.SaveChangesAsync();
                    await component.UpdateAsync(m => { m.Content = "🔥 資料已清空。"; m.Components = null; });
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

            var existingUsers = await db.UserPoints.Where(u => u.GuildId == guildId).Select(u => u.RobloxUsername.ToLower()).ToListAsync();
            var existingSet = new HashSet<string>(existingUsers);
            var validActiveUsers = new HashSet<string>();

            while (hasMore)
            {
                string url = $"https://groups.roblox.com/v1/groups/{groupId}/users?limit=100";
                if (!string.IsNullOrEmpty(cursor)) url += $"&cursor={cursor}";

                var res = await _http.GetAsync(url);
                if (!res.IsSuccessStatusCode)
                {
                    Console.WriteLine("⚠️ Roblox API 請求失敗或已達限制，中斷同步以保護資料庫。");
                    return (addedCount, removedCount); // 💡 修改為回傳兩個值
                }

                var json = await res.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                var data = root.GetProperty("data");
                foreach (var item in data.EnumerateArray())
                {
                    string username = item.GetProperty("user").GetProperty("username").GetString();
                    int rank = item.GetProperty("role").GetProperty("rank").GetInt32();

                    if (rank <= maxRank && rank >= minRank)
                    {
                        validActiveUsers.Add(username.ToLower());

                        if (!existingSet.Contains(username.ToLower()))
                        {
                            db.UserPoints.Add(new UserPoint { GuildId = guildId, RobloxUsername = username, Points = 0 });
                            existingSet.Add(username.ToLower());
                            addedCount++;
                        }
                    }
                }

                if (root.TryGetProperty("nextPageCursor", out var nextCursor) && nextCursor.ValueKind == JsonValueKind.String)
                    cursor = nextCursor.GetString();
                else
                    hasMore = false;
            }

            var usersToDelete = await db.UserPoints.Where(u => u.GuildId == guildId).ToListAsync();
            var finalDeleteList = usersToDelete.Where(u => !validActiveUsers.Contains(u.RobloxUsername.ToLower())).ToList();

            if (finalDeleteList.Any())
            {
                db.UserPoints.RemoveRange(finalDeleteList);
                removedCount = finalDeleteList.Count;
                Console.WriteLine($"🗑️ 已自動清理 {removedCount} 位不符合權重或退群的成員。");
            }

            if (addedCount > 0 || removedCount > 0) await db.SaveChangesAsync();

            return (addedCount, removedCount); // 💡 修改為回傳兩個值
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

        private async Task BroadcastToAdminChannelsAsync(ulong sourceGuildId, string message)
        {
            using var db = new BotDbContext(_configuration);
            var sourceConfig = await db.Configs.FindAsync(sourceGuildId);
            if (sourceConfig == null || string.IsNullOrEmpty(sourceConfig.ServerCode)) return;

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
    }

    public class BotDbContext : DbContext
    {
        private readonly IConfiguration _config;
        public BotDbContext(IConfiguration config) { _config = config; }

        public DbSet<UserPoint> UserPoints { get; set; }
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

    public class UserPoint { [Key] public int Id { get; set; } public ulong GuildId { get; set; } public string RobloxUsername { get; set; } public int Points { get; set; } }
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
    }
    // 👇 新增這個類別：記錄哪個頻道負責監聽哪些伺服器編號
    public class AdminChannelConfig
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public ulong ChannelId { get; set; }
        public ulong GuildId { get; set; }
        public string TargetServerCodes { get; set; } // 存放要監控的編號，例如 "A1B2C,X9Y8Z"
    }
}