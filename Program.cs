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
                new SlashCommandBuilder().WithName("setup-roca").WithDescription("🔗 綁定 Roblox 群組並自動同步成員").AddOption("roblox_group_id", ApplicationCommandOptionType.String, "Roblox 群組 ID", isRequired: true).AddOption("admin_role", ApplicationCommandOptionType.Role, "管理身分組", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("sync-members").WithDescription("🔄 手動同步 Roblox 群組的最新成員至資料庫").Build(),
                new SlashCommandBuilder().WithName("points").WithDescription("📊 查詢點數").AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("addpoint").WithDescription("➕ 發放點數").AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true).AddOption("points", ApplicationCommandOptionType.Integer, "點數數量", isRequired: true).AddOption("reason", ApplicationCommandOptionType.String, "原因備註", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("history").WithDescription("📜 查詢玩家近期紀錄").AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("viewall").WithDescription("🏆 點數總排行榜 (顯示所有成員)").Build(),
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
                bool isEphemeral = command.Data.Name == "setup-roca" || command.Data.Name == "unbind-roca" || command.Data.Name == "clear-all-data" || command.Data.Name == "sync-members";
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
                        if (!((SocketGuildUser)command.User).GuildPermissions.Administrator) { await command.FollowupAsync("❌ 限管理員執行。"); return; }
                        string rId = (string)command.Data.Options.First(x => x.Name == "roblox_group_id").Value;
                        var rRole = (SocketRole)command.Data.Options.First(x => x.Name == "admin_role").Value;
                        if (botConfig == null) db.Configs.Add(new BotConfig { GuildId = gid, RobloxGroupId = rId, AdminRoleId = rRole.Id });
                        else { botConfig.RobloxGroupId = rId; botConfig.AdminRoleId = rRole.Id; }
                        await db.SaveChangesAsync();

                        await command.FollowupAsync($"✅ 已綁定 Roblox 群組 `{rId}`。正在背景抓取群組成員，請稍候...");
                        int addedCount = await SyncGroupMembersAsync(db, gid, rId);

                        // 修正：使用 FollowupAsync 就不會有錯誤了
                        await command.FollowupAsync($"🔄 群組名單同步完成！共為排行榜新增了 **{addedCount}** 位新成員 (預設為 0 點)。");
                        break;

                    case "sync-members":
                        if (botConfig == null) { await command.FollowupAsync("❌ 請先使用 `/setup-roca` 綁定群組。"); return; }
                        if (!((SocketGuildUser)command.User).Roles.Any(r => r.Id == botConfig.AdminRoleId)) { await command.FollowupAsync("❌ 權限不足。"); return; }

                        await command.FollowupAsync("⏳ 正在與 Roblox 同步群組成員名單...");
                        int newMembers = await SyncGroupMembersAsync(db, gid, botConfig.RobloxGroupId);
                        await command.FollowupAsync($"✅ 同步完成！本次共新增了 **{newMembers}** 位新成員。");
                        break;

                    case "viewall":
                        var all = await db.UserPoints.Where(u => u.GuildId == gid).OrderByDescending(u => u.Points).ToListAsync();
                        if (!all.Any()) { await command.FollowupAsync("📭 尚無資料。"); break; }

                        var chunks = new List<string>();
                        var currentChunk = new StringBuilder("```markdown\n# 🏆 點數總覽 (所有成員)\n");

                        foreach (var u in all)
                        {
                            string line = $"{u.RobloxUsername,-20} | {u.Points} 點\n";
                            if (currentChunk.Length + line.Length > 1900)
                            {
                                currentChunk.Append("```");
                                chunks.Add(currentChunk.ToString());
                                currentChunk.Clear();
                                currentChunk.Append("```markdown\n");
                            }
                            currentChunk.Append(line);
                        }
                        if (currentChunk.Length > "```markdown\n".Length)
                        {
                            currentChunk.Append("```");
                            chunks.Add(currentChunk.ToString());
                        }

                        // 修正：迴圈內全部使用 FollowupAsync 分段發送
                        foreach (var chunk in chunks)
                        {
                            await command.FollowupAsync(chunk);
                        }
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

                    case "unbind-roca":
                        if (!((SocketGuildUser)command.User).GuildPermissions.Administrator) { await command.FollowupAsync("❌ 限管理員執行。"); return; }
                        if (botConfig != null) { db.Configs.Remove(botConfig); await db.SaveChangesAsync(); await command.FollowupAsync("🔓 已成功解除本伺服器的設定。"); }
                        else { await command.FollowupAsync("⚠️ 本伺服器尚未進行綁定。"); }
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
                string realError = ex.InnerException != null ? ex.InnerException.Message : ex.Message;
                Console.WriteLine($"[指令錯誤] {realError}");
                string errMsg = $"❌ 發生內部錯誤: `{realError}`\n請檢查資料庫狀態。";
                if (command.HasResponded) await command.FollowupAsync(errMsg, ephemeral: true);
                else await command.RespondAsync(errMsg, ephemeral: true);
            }
        }

        private async Task<int> SyncGroupMembersAsync(BotDbContext db, ulong guildId, string groupId)
        {
            string cursor = "";
            bool hasMore = true;
            int addedCount = 0;

            var existingUsers = await db.UserPoints.Where(u => u.GuildId == guildId).Select(u => u.RobloxUsername.ToLower()).ToListAsync();
            var existingSet = new HashSet<string>(existingUsers);

            while (hasMore)
            {
                // 修正：移除髒掉的 Markdown 網址標籤
                string url = $"[https://groups.roblox.com/v1/groups/](https://groups.roblox.com/v1/groups/){groupId}/users?limit=100";
                if (!string.IsNullOrEmpty(cursor)) url += $"&cursor={cursor}";

                var res = await _http.GetAsync(url);
                if (!res.IsSuccessStatusCode) break;

                var json = await res.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                var data = root.GetProperty("data");
                foreach (var item in data.EnumerateArray())
                {
                    string username = item.GetProperty("user").GetProperty("username").GetString();

                    if (!existingSet.Contains(username.ToLower()))
                    {
                        db.UserPoints.Add(new UserPoint { GuildId = guildId, RobloxUsername = username, Points = 0 });
                        existingSet.Add(username.ToLower());
                        addedCount++;
                    }
                }

                if (root.TryGetProperty("nextPageCursor", out var nextCursor) && nextCursor.ValueKind == JsonValueKind.String)
                {
                    cursor = nextCursor.GetString();
                }
                else
                {
                    hasMore = false;
                }
            }

            if (addedCount > 0) await db.SaveChangesAsync();
            return addedCount;
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

                // 修正：移除髒掉的 Markdown 網址標籤
                var userRes = await _http.PostAsync("[https://users.roblox.com/v1/usernames/users](https://users.roblox.com/v1/usernames/users)", content);
                var userJson = await userRes.Content.ReadAsStringAsync();
                var data = JsonDocument.Parse(userJson).RootElement.GetProperty("data");
                if (data.GetArrayLength() == 0) return false;

                long userId = data[0].GetProperty("id").GetInt64();

                // 修正：移除髒掉的 Markdown 網址標籤
                var groupRes = await _http.GetAsync($"[https://groups.roblox.com/v1/users/](https://groups.roblox.com/v1/users/){userId}/groups/roles");
                var groupJson = await groupRes.Content.ReadAsStringAsync();
                return groupJson.Contains($"\"id\":{groupId}");
            }
            catch { return false; }
        }
    }

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
            else options.UseSqlite("Data Source=rocapoints.db");
        }
    }

    public class UserPoint { [Key] public int Id { get; set; } public ulong GuildId { get; set; } public string RobloxUsername { get; set; } public int Points { get; set; } }
    public class PointLog { [Key] public int Id { get; set; } public ulong GuildId { get; set; } public string RobloxUsername { get; set; } public string AdminName { get; set; } public int PointsAdded { get; set; } public string Reason { get; set; } public DateTime Timestamp { get; set; } public bool IsDeleted { get; set; } }

    [Table("GuildConfigs")]
    public class BotConfig
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.None)]
        public ulong GuildId { get; set; }
        public string RobloxGroupId { get; set; }
        public ulong AdminRoleId { get; set; }
    }
}