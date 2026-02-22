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
            app.MapGet("/", () => "✅ ROCA Point Bot (SQL Server 版) 正在運行中！");
            app.Run();
        }

        // 強制台北時區輔助函數
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

            // 1. 先讓機器人登入並上線！
            await _client.LoginAsync(TokenType.Bot, _discordToken);
            await _client.StartAsync();

            // 2. 在背景檢查並建立資料庫
            _ = Task.Run(async () =>
            {
                try
                {
                    using var db = new BotDbContext(_configuration);

                    // ⚠️ 【一次性重建指令】先把舊的、錯誤的表格通通刪除
                    await db.Database.ExecuteSqlRawAsync(
                        "DROP TABLE IF EXISTS UserPoints; " +
                        "DROP TABLE IF EXISTS PointLogs; " +
                        "DROP TABLE IF EXISTS Configs; " +
                        "DROP TABLE IF EXISTS GuildConfigs;"
                    );

                    // 重新建立包含正確設定的所有新表格！
                    await db.Database.EnsureCreatedAsync();

                    Console.WriteLine("✅ 資料庫表格已成功清除並完美重建！");
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
                new SlashCommandBuilder().WithName("setup-roca").WithDescription("🔗 綁定 Roblox 群組與管理身分組").AddOption("roblox_group_id", ApplicationCommandOptionType.String, "Roblox 群組 ID", isRequired: true).AddOption("admin_role", ApplicationCommandOptionType.Role, "管理身分組", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("points").WithDescription("📊 查詢點數").AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("addpoint").WithDescription("➕ 發放點數").AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true).AddOption("points", ApplicationCommandOptionType.Integer, "點數數量", isRequired: true).AddOption("reason", ApplicationCommandOptionType.String, "原因備註", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("history").WithDescription("📜 查詢玩家近期紀錄").AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("viewall").WithDescription("🏆 點數總排行榜").Build(),
                new SlashCommandBuilder().WithName("del-record").WithDescription("🚨 刪除單筆紀錄").AddOption("id", ApplicationCommandOptionType.Integer, "紀錄 ID", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("unbind-roca").WithDescription("🔓 解除綁定設定").Build(),
                new SlashCommandBuilder().WithName("clear-all-data").WithDescription("🔥 【極度危險】刪除本伺服器所有資料").Build(),
                new SlashCommandBuilder().WithName("daily-records").WithDescription("📅 查詢指定日期的所有點數紀錄").AddOption("date", ApplicationCommandOptionType.String, "格式: YYYY-MM-DD", isRequired: true).Build()
            };
            try { await _client.BulkOverwriteGlobalApplicationCommandsAsync(commands.ToArray()); } catch (Exception ex) { Console.WriteLine(ex.Message); }
        }

        private async Task HandleSlashCommandAsync(SocketSlashCommand command)
        {
            if (!(command.Channel is SocketGuildChannel guildChannel)) return;
            ulong gid = guildChannel.Guild.Id;

            try
            {
                // 1. 預先判斷指令是否需要「隱藏回覆」 (只有自己看得到)
                bool isEphemeral = command.Data.Name == "setup-roca" || command.Data.Name == "unbind-roca" || command.Data.Name == "clear-all-data";

                // 2. 收到指令的第一時間先 Defer，爭取 15 分鐘的處理時間！
                await command.DeferAsync(ephemeral: isEphemeral);

                using var db = new BotDbContext(_configuration);

                // 3. 測試資料庫連線是否暢通
                if (!await db.Database.CanConnectAsync())
                {
                    await command.FollowupAsync("❌ 無法連線到資料庫，請檢查遠端資料庫狀態或防火牆 IP 阻擋設定。");
                    return;
                }

                var botConfig = await db.Configs.FindAsync(gid);

                switch (command.Data.Name)
                {
                    case "setup-roca":
                        if (!((SocketGuildUser)command.User).GuildPermissions.Administrator) { await command.FollowupAsync("❌ 限管理員執行。"); return; }
                        string rId = (string)command.Data.Options.First(x => x.Name == "roblox_group_id").Value;
                        var rRole = (SocketRole)command.Data.Options.First(x => x.Name == "admin_role").Value;
                        if (botConfig == null) db.Configs.Add(new BotConfig { GuildId = gid, RobloxGroupId = rId, AdminRoleId = rRole.Id });
                        else { botConfig.RobloxGroupId = rId; botConfig.AdminRoleId = rRole.Id; }
                        await db.SaveChangesAsync();
                        await command.FollowupAsync($"✅ 已綁定 Roblox 群組 `{rId}`。");
                        break;

                    case "addpoint":
                        if (botConfig == null) { await command.FollowupAsync("❌ 未設定。請先使用 /setup-roca"); break; }
                        var exec = (SocketGuildUser)command.User;
                        if (!exec.Roles.Any(r => r.Id == botConfig.AdminRoleId) && exec.Id != guildChannel.Guild.OwnerId) { await command.FollowupAsync("❌ 權限不足。"); return; }
                        var user = (SocketGuildUser)command.Data.Options.First(x => x.Name == "user").Value;
                        int pts = Convert.ToInt32((long)command.Data.Options.First(x => x.Name == "points").Value);
                        string reason = (string)command.Data.Options.First(x => x.Name == "reason").Value;
                        string name = user.Nickname ?? user.Username;

                        if (!await VerifyUserInRobloxGroup(name, botConfig.RobloxGroupId)) { await command.FollowupAsync($"❌ 玩家 `{name}` 不在指定的 Roblox 群組內，或名稱不相符。"); break; }

                        var rec = await db.UserPoints.FirstOrDefaultAsync(u => u.GuildId == gid && u.RobloxUsername.ToLower() == name.ToLower());
                        if (rec == null) { rec = new UserPoint { GuildId = gid, RobloxUsername = name, Points = 0 }; db.UserPoints.Add(rec); }
                        rec.Points += pts;
                        db.PointLogs.Add(new PointLog { GuildId = gid, RobloxUsername = name, AdminName = command.User.Username, PointsAdded = pts, Reason = reason, Timestamp = Program.GetTaipeiTime() });
                        await db.SaveChangesAsync();
                        await command.FollowupAsync($"✅ 已發放 {pts} 點給 {name}。");
                        break;

                    case "history":
                        var hUser = (SocketGuildUser)command.Data.Options.First().Value;
                        string hName = hUser.Nickname ?? hUser.Username;
                        var logs = await db.PointLogs.Where(l => l.GuildId == gid && l.RobloxUsername.ToLower() == hName.ToLower() && !l.IsDeleted).OrderByDescending(l => l.Timestamp).Take(10).ToListAsync();
                        if (!logs.Any()) { await command.FollowupAsync("📭 查無紀錄。"); break; }
                        var sb = new StringBuilder($"📜 **{hName}** 的個人紀錄：\n");
                        foreach (var l in logs) sb.AppendLine($"`#{l.Id}` | {l.Timestamp:MM/dd HH:mm} | **+{l.PointsAdded}** | {l.Reason}");
                        await command.FollowupAsync(sb.ToString());
                        break;

                    case "del-record":
                        if (botConfig == null) { await command.FollowupAsync("⚠️ 請先設定。"); return; }
                        if (!((SocketGuildUser)command.User).Roles.Any(r => r.Id == botConfig.AdminRoleId)) { await command.FollowupAsync("❌ 權限不足。"); return; }
                        int logId = Convert.ToInt32((long)command.Data.Options.First().Value);
                        var targetLog = await db.PointLogs.FirstOrDefaultAsync(l => l.Id == logId && l.GuildId == gid);
                        if (targetLog == null || targetLog.IsDeleted) { await command.FollowupAsync("❌ 找不到該紀錄。"); break; }
                        var uPoint = await db.UserPoints.FirstOrDefaultAsync(u => u.GuildId == gid && u.RobloxUsername.ToLower() == targetLog.RobloxUsername.ToLower());
                        if (uPoint != null) uPoint.Points = Math.Max(0, uPoint.Points - targetLog.PointsAdded);
                        targetLog.IsDeleted = true;
                        await db.SaveChangesAsync();
                        await command.FollowupAsync($"🗑️ 紀錄 #{logId} 已刪除，點數已扣回。");
                        break;

                    case "daily-records":
                        string dateStr = (string)command.Data.Options.First().Value;
                        if (!DateTime.TryParse(dateStr, out DateTime dt)) { await command.FollowupAsync("❌ 日期格式錯誤。"); break; }
                        var dLogs = await db.PointLogs.Where(l => l.GuildId == gid && l.Timestamp.Date == dt.Date && !l.IsDeleted).ToListAsync();
                        if (!dLogs.Any()) { await command.FollowupAsync("📭 該日無紀錄。"); break; }
                        var dSb = new StringBuilder($"# 📅 {dt:yyyy-MM-dd} 紀錄清單\n```markdown\n");
                        foreach (var l in dLogs) dSb.AppendLine($"{l.Timestamp:HH:mm} | {l.RobloxUsername,-15} | +{l.PointsAdded}");
                        dSb.Append("```");
                        await command.FollowupAsync(dSb.ToString());
                        break;

                    case "clear-all-data":
                        if (!((SocketGuildUser)command.User).GuildPermissions.Administrator) { await command.FollowupAsync("❌ 限管理員。"); return; }
                        var btns = new ComponentBuilder().WithButton("確認刪除 (1/2)", $"clear_step1_{gid}", ButtonStyle.Danger).WithButton("取消", $"clear_cancel_{gid}", ButtonStyle.Secondary);
                        await command.FollowupAsync("⚠️ **【第一層確認】** 確定要清空所有資料嗎？", components: btns.Build());
                        break;

                    case "viewall":
                        var all = await db.UserPoints.Where(u => u.GuildId == gid).OrderByDescending(u => u.Points).Take(30).ToListAsync();
                        if (!all.Any()) { await command.FollowupAsync("📭 尚無資料。"); break; }
                        var lb = new StringBuilder("```markdown\n# 🏆 點數總覽 (台北時間)\n");
                        foreach (var u in all) lb.AppendLine($"{u.RobloxUsername,-20} | {u.Points} 點");
                        lb.Append("```");
                        await command.FollowupAsync(lb.ToString());
                        break;

                    case "unbind-roca":
                        if (!((SocketGuildUser)command.User).GuildPermissions.Administrator) { await command.FollowupAsync("❌ 限管理員執行。"); return; }
                        if (botConfig != null)
                        {
                            db.Configs.Remove(botConfig);
                            await db.SaveChangesAsync();
                            await command.FollowupAsync("🔓 已成功解除本伺服器的 Roblox 群組與身分組綁定設定。");
                        }
                        else { await command.FollowupAsync("⚠️ 本伺服器尚未進行綁定，無需解除。"); }
                        break;

                    case "points":
                        var targetUser = (SocketGuildUser)command.Data.Options.First(x => x.Name == "user").Value;
                        string targetName = targetUser.Nickname ?? targetUser.Username;
                        var userPoint = await db.UserPoints.FirstOrDefaultAsync(u => u.GuildId == gid && u.RobloxUsername.ToLower() == targetName.ToLower());
                        int currentPoints = userPoint != null ? userPoint.Points : 0;
                        await command.FollowupAsync($"📊 **{targetName}** 目前擁有 **{currentPoints}** 點。");
                        break;
                }
            }
            catch (Exception ex)
            {
                // 抓取最深層的真實錯誤訊息
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

        private async Task<bool> VerifyUserInRobloxGroup(string username, string groupId)
        {
            try
            {
                var userReq = new { usernames = new[] { username }, excludeBannedUsers = true };
                var content = new StringContent(JsonSerializer.Serialize(userReq), Encoding.UTF8, "application/json");
                var userRes = await _http.PostAsync("https://users.roblox.com/v1/usernames/users", content);
                var userJson = await userRes.Content.ReadAsStringAsync();
                var data = JsonDocument.Parse(userJson).RootElement.GetProperty("data");
                if (data.GetArrayLength() == 0) return false;
                long userId = data[0].GetProperty("id").GetInt64();
                var groupRes = await _http.GetAsync($"https://groups.roblox.com/v1/users/{userId}/groups/roles");
                var groupJson = await groupRes.Content.ReadAsStringAsync();
                return groupJson.Contains($"\"id\":{groupId}");
            }
            catch { return false; }
        }
    }

    // ==========================================
    // SQL Server 資料庫設定檔
    // ==========================================
    public class BotDbContext : DbContext
    {
        private readonly IConfiguration _config;
        public BotDbContext(IConfiguration config) { _config = config; }

        public DbSet<UserPoint> UserPoints { get; set; }
        public DbSet<PointLog> PointLogs { get; set; }
        public DbSet<BotConfig> Configs { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder options)
        {
            string conn = _config["DbConnection"];
            if (!string.IsNullOrEmpty(conn)) options.UseSqlServer(conn);
            else options.UseSqlite("Data Source=rocapoints.db"); // 備援用
        }
    }

    public class UserPoint { [Key] public int Id { get; set; } public ulong GuildId { get; set; } public string RobloxUsername { get; set; } public int Points { get; set; } }
    public class PointLog { [Key] public int Id { get; set; } public ulong GuildId { get; set; } public string RobloxUsername { get; set; } public string AdminName { get; set; } public int PointsAdded { get; set; } public string Reason { get; set; } public DateTime Timestamp { get; set; } public bool IsDeleted { get; set; } }
    // 加入這個標籤，讓程式在資料庫建一個新表格，避開舊的錯誤表格
    [Table("GuildConfigs")]
    public class BotConfig
    {
        [Key]
        // 加上這行！這會關閉資料庫的「自動遞增」功能，允許我們寫入 Discord ID
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public ulong GuildId { get; set; }

        public string RobloxGroupId { get; set; }
        public ulong AdminRoleId { get; set; }
    }
}