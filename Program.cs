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

namespace ROCAPointBot
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);
            builder.Services.AddHostedService<DiscordBotService>();
            var app = builder.Build();
            app.MapGet("/", () => "✅ ROCA Point Bot 專業版 (雙重確認刪除機制已啟動)");
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

        public DiscordBotService(IConfiguration config)
        {
            _discordToken = config["DiscordToken"];
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            using (var db = new BotDbContext()) { await db.Database.EnsureCreatedAsync(); }

            var config = new DiscordSocketConfig { GatewayIntents = GatewayIntents.AllUnprivileged | GatewayIntents.GuildMembers };
            _client = new DiscordSocketClient(config);
            _client.Ready += RegisterCommandsAsync;
            _client.SlashCommandExecuted += HandleSlashCommandAsync;
            _client.InteractionCreated += HandleInteractionAsync; // 處理按鈕

            await _client.LoginAsync(TokenType.Bot, _discordToken);
            await _client.StartAsync();
            await Task.Delay(Timeout.Infinite, stoppingToken);
        }

        private async Task RegisterCommandsAsync()
        {
            var commands = new List<ApplicationCommandProperties>
            {
                new SlashCommandBuilder().WithName("setup-roca").WithDescription("🔗 綁定 Roblox 群組與管理身分組").AddOption("roblox_group_id", ApplicationCommandOptionType.String, "Roblox 群組 ID", isRequired: true).AddOption("admin_role", ApplicationCommandOptionType.Role, "管理身分組", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("points").WithDescription("📊 查詢點數").AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("addpoint").WithDescription("➕ 發放點數").AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true).AddOption("points", ApplicationCommandOptionType.Integer, "點數數量", isRequired: true).AddOption("reason", ApplicationCommandOptionType.String, "原因備註", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("history").WithDescription("📜 查詢近期紀錄").AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("viewall").WithDescription("🏆 點數總排行榜").Build(),
                new SlashCommandBuilder().WithName("del-record").WithDescription("🚨 刪除單筆紀錄").AddOption("id", ApplicationCommandOptionType.Integer, "紀錄 ID", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("unbind-roca").WithDescription("🔓 解除綁定設定").Build(),
                // --- 新增：全資料刪除指令 ---
                new SlashCommandBuilder().WithName("clear-all-data").WithDescription("🔥 【極度危險】刪除本伺服器所有成員點數與紀錄").Build()
            };
            try { await _client.BulkOverwriteGlobalApplicationCommandsAsync(commands.ToArray()); } catch (Exception ex) { Console.WriteLine(ex.Message); }
        }

        private async Task HandleSlashCommandAsync(SocketSlashCommand command)
        {
            if (!(command.Channel is SocketGuildChannel guildChannel)) return;
            ulong gid = guildChannel.Guild.Id;

            try
            {
                switch (command.Data.Name)
                {
                    case "clear-all-data":
                        // 僅限管理員
                        if (!((SocketGuildUser)command.User).GuildPermissions.Administrator)
                        {
                            await command.RespondAsync("❌ 僅限伺服器管理員執行此項危險操作。", ephemeral: true);
                            return;
                        }

                        var btn1 = new ComponentBuilder()
                            .WithButton("確認刪除 (1/2)", $"clear_step1_{gid}", ButtonStyle.Danger)
                            .WithButton("取消", $"clear_cancel_{gid}", ButtonStyle.Secondary);

                        await command.RespondAsync("⚠️ **【第一層確認】**\n您即將刪除本伺服器「所有成員」的點數資料與歷史紀錄。\n此操作**不可逆**，請問確定要繼續嗎？", components: btn1.Build());
                        break;

                    // ... 其他指令邏輯 (維持不變) ...
                    case "setup-roca":
                        await command.DeferAsync();
                        using (var db = new BotDbContext())
                        {
                            string rId = (string)command.Data.Options.First(x => x.Name == "roblox_group_id").Value;
                            var rRole = (SocketRole)command.Data.Options.First(x => x.Name == "admin_role").Value;
                            var cfg = await db.Configs.FindAsync(gid);
                            if (cfg == null) db.Configs.Add(new BotConfig { GuildId = gid, RobloxGroupId = rId, AdminRoleId = rRole.Id });
                            else { cfg.RobloxGroupId = rId; cfg.AdminRoleId = rRole.Id; }
                            await db.SaveChangesAsync();
                            await command.FollowupAsync($"✅ 已綁定 Roblox 群組 `{rId}`。");
                        }
                        break;

                    case "addpoint":
                        await command.DeferAsync();
                        using (var db = new BotDbContext())
                        {
                            var cfg = await db.Configs.FindAsync(gid);
                            if (cfg == null) { await command.FollowupAsync("❌ 未設定。"); break; }
                            var user = (SocketGuildUser)command.Data.Options.First(x => x.Name == "user").Value;
                            int pts = Convert.ToInt32((long)command.Data.Options.First(x => x.Name == "points").Value);
                            string reason = (string)command.Data.Options.First(x => x.Name == "reason").Value;
                            string name = user.Nickname ?? user.Username;

                            if (!await VerifyUserInRobloxGroup(name, cfg.RobloxGroupId)) { await command.FollowupAsync("❌ 不在群組內。"); break; }

                            var rec = await db.UserPoints.FirstOrDefaultAsync(u => u.GuildId == gid && u.RobloxUsername.ToLower() == name.ToLower());
                            if (rec == null) { rec = new UserPoint { GuildId = gid, RobloxUsername = name, Points = 0 }; db.UserPoints.Add(rec); }
                            rec.Points += pts;
                            db.PointLogs.Add(new PointLog { GuildId = gid, RobloxUsername = name, AdminName = command.User.Username, PointsAdded = pts, Reason = reason, Timestamp = Program.GetTaipeiTime() });
                            await db.SaveChangesAsync();
                            await command.FollowupAsync($"✅ 已發放 {pts} 點給 {name}。");
                        }
                        break;

                    case "history":
                        await command.DeferAsync();
                        using (var db = new BotDbContext())
                        {
                            var user = (SocketGuildUser)command.Data.Options.First().Value;
                            string name = user.Nickname ?? user.Username;
                            var logs = await db.PointLogs.Where(l => l.GuildId == gid && l.RobloxUsername.ToLower() == name.ToLower() && !l.IsDeleted).OrderByDescending(l => l.Timestamp).Take(10).ToListAsync();
                            if (!logs.Any()) { await command.FollowupAsync("📭 查無紀錄。"); break; }
                            var sb = new StringBuilder($"📜 **{name}** 的紀錄：\n");
                            foreach (var l in logs) sb.AppendLine($"`#{l.Id}` | {l.Timestamp:MM/dd} | **+{l.PointsAdded}** | {l.Reason}");
                            await command.FollowupAsync(sb.ToString());
                        }
                        break;

                    case "viewall":
                        await command.DeferAsync();
                        using (var db = new BotDbContext())
                        {
                            var all = await db.UserPoints.Where(u => u.GuildId == gid).OrderByDescending(u => u.Points).ToListAsync();
                            if (!all.Any()) { await command.FollowupAsync("📭 尚無資料。"); break; }
                            var lb = new StringBuilder("```markdown\n# 🏆 排行榜\n");
                            foreach (var u in all) lb.AppendLine($"{u.RobloxUsername,-20} | {u.Points} 點");
                            lb.Append("```");
                            await command.FollowupAsync(lb.ToString());
                        }
                        break;
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }

        // ==========================================
        // 處理雙重確認按鈕互動
        // ==========================================
        private async Task HandleInteractionAsync(SocketInteraction interaction)
        {
            if (interaction is SocketMessageComponent component)
            {
                string id = component.Data.CustomId;
                ulong gid = (ulong)component.GuildId;

                // 1. 取消動作
                if (id.StartsWith("clear_cancel_"))
                {
                    await component.UpdateAsync(m => { m.Content = "❌ 指令已取消，資料未受影響。"; m.Components = null; });
                    return;
                }

                // 2. 第一層確認完畢，切換到第二層
                if (id.StartsWith("clear_step1_"))
                {
                    var btn2 = new ComponentBuilder()
                        .WithButton("⚠️ 最終確認刪除 (2/2)", $"clear_step2_{gid}", ButtonStyle.Danger)
                        .WithButton("取消", $"clear_cancel_{gid}", ButtonStyle.Secondary);

                    await component.UpdateAsync(m => {
                        m.Content = "🚨 **【最終警告】**\n按下按鈕後，**本伺服器的所有點數紀錄將會瞬間清空**，無法復原！\n確認執行請點擊下方紅按鈕。";
                        m.Components = btn2.Build();
                    });
                    return;
                }

                // 3. 第二層確認，正式執行刪除
                if (id.StartsWith("clear_step2_"))
                {
                    using var db = new BotDbContext();
                    // 刪除點數紀錄
                    var points = db.UserPoints.Where(u => u.GuildId == gid);
                    db.UserPoints.RemoveRange(points);
                    // 標記日誌為刪除
                    var logs = db.PointLogs.Where(l => l.GuildId == gid);
                    foreach (var l in logs) { l.IsDeleted = true; l.Reason = "[伺服器重置] " + l.Reason; }

                    await db.SaveChangesAsync();

                    await component.UpdateAsync(m => {
                        m.Content = $"🔥 **資料已清空！**\n本伺服器所有成員點數已重置為 0，且所有歷史紀錄已歸檔刪除。\n操作者：{component.User.Username}\n時間：{Program.GetTaipeiTime():yyyy/MM/dd HH:mm:ss}";
                        m.Components = null;
                    });
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

    public class BotDbContext : DbContext
    {
        public DbSet<UserPoint> UserPoints { get; set; }
        public DbSet<PointLog> PointLogs { get; set; }
        public DbSet<BotConfig> Configs { get; set; }
        protected override void OnConfiguring(DbContextOptionsBuilder options) => options.UseSqlite("Data Source=rocapoints.db");
    }
    public class UserPoint { [Key] public int Id { get; set; } public ulong GuildId { get; set; } public string RobloxUsername { get; set; } public int Points { get; set; } }
    public class PointLog { [Key] public int Id { get; set; } public ulong GuildId { get; set; } public string RobloxUsername { get; set; } public string AdminName { get; set; } public int PointsAdded { get; set; } public string Reason { get; set; } public DateTime Timestamp { get; set; } public bool IsDeleted { get; set; } }
    public class BotConfig { [Key] public ulong GuildId { get; set; } public string RobloxGroupId { get; set; } public ulong AdminRoleId { get; set; } }
}