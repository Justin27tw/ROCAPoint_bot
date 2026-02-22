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
            app.MapGet("/", () => "✅ ROCA Point Bot 正在台北時區運行中！");
            app.Run();
        }

        // --- 台北時區轉換輔助函數 ---
        public static DateTime GetTaipeiTime()
        {
            DateTime utcNow = DateTime.UtcNow;
            try
            {
                // 適用於 Windows 伺服器 (runasp.net)
                return TimeZoneInfo.ConvertTimeFromUtc(utcNow, TimeZoneInfo.FindSystemTimeZoneById("Taipei Standard Time"));
            }
            catch
            {
                // 適用於 Linux 伺服器或備用方案
                return utcNow.AddHours(8);
            }
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
            using (var db = new BotDbContext())
            {
                await db.Database.EnsureCreatedAsync();
            }

            var config = new DiscordSocketConfig
            {
                GatewayIntents = GatewayIntents.AllUnprivileged | GatewayIntents.GuildMembers
            };

            _client = new DiscordSocketClient(config);
            _client.Ready += RegisterCommandsAsync;
            _client.SlashCommandExecuted += HandleSlashCommandAsync;
            _client.InteractionCreated += HandleInteractionAsync;

            await _client.LoginAsync(TokenType.Bot, _discordToken);
            await _client.StartAsync();

            await Task.Delay(Timeout.Infinite, stoppingToken);
        }

        private async Task RegisterCommandsAsync()
        {
            var commands = new List<ApplicationCommandProperties>
            {
                // 核心：綁定指令 (Discord 群組對應一個 Roblox 群組 + 權限身分組)
                new SlashCommandBuilder()
                    .WithName("setup-roca")
                    .WithDescription("🔗 綁定本伺服器的 Roblox 群組與管理權限")
                    .AddOption("roblox_group_id", ApplicationCommandOptionType.String, "Roblox 群組 ID", isRequired: true)
                    .AddOption("admin_role", ApplicationCommandOptionType.Role, "允許操作點數的 Discord 身分組", isRequired: true)
                    .Build(),

                new SlashCommandBuilder()
                    .WithName("points")
                    .WithDescription("📊 查詢玩家點數")
                    .AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true)
                    .Build(),

                new SlashCommandBuilder()
                    .WithName("addpoint")
                    .WithDescription("➕ 發放點數 (限指定管理身分組使用)")
                    .AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true)
                    .AddOption("points", ApplicationCommandOptionType.Integer, "點數數量", isRequired: true)
                    .AddOption("reason", ApplicationCommandOptionType.String, "原因備註", isRequired: true)
                    .Build(),

                new SlashCommandBuilder()
                    .WithName("viewall")
                    .WithDescription("🏆 查看本伺服器所有成員點數總覽")
                    .Build(),

                new SlashCommandBuilder()
                    .WithName("history")
                    .WithDescription("📜 查詢玩家近期紀錄")
                    .AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true)
                    .Build()
            };

            try { await _client.BulkOverwriteGlobalApplicationCommandsAsync(commands.ToArray()); }
            catch (Exception ex) { Console.WriteLine($"註冊錯誤: {ex.Message}"); }
        }

        private async Task HandleSlashCommandAsync(SocketSlashCommand command)
        {
            if (!(command.Channel is SocketGuildChannel guildChannel)) return;
            await command.DeferAsync();

            ulong currentGuildId = guildChannel.Guild.Id;
            using var db = new BotDbContext();

            // 讀取該伺服器的專屬綁定設定
            var botConfig = await db.Configs.FirstOrDefaultAsync(c => c.GuildId == currentGuildId);

            try
            {
                switch (command.Data.Name)
                {
                    case "setup-roca":
                        // 僅限伺服器擁有人或管理員設定
                        if (!((SocketGuildUser)command.User).GuildPermissions.Administrator)
                        {
                            await command.FollowupAsync("❌ 只有伺服器管理員可以執行初始化設定。");
                            return;
                        }

                        string rGroupId = (string)command.Data.Options.First(x => x.Name == "roblox_group_id").Value;
                        var rRole = (SocketRole)command.Data.Options.First(x => x.Name == "admin_role").Value;

                        if (botConfig == null)
                        {
                            botConfig = new BotConfig { GuildId = currentGuildId, RobloxGroupId = rGroupId, AdminRoleId = rRole.Id };
                            db.Configs.Add(botConfig);
                        }
                        else
                        {
                            botConfig.RobloxGroupId = rGroupId;
                            botConfig.AdminRoleId = rRole.Id;
                        }
                        await db.SaveChangesAsync();
                        await command.FollowupAsync($"✅ 設定完成！\n🔗 綁定 Roblox 群組：`{rGroupId}`\n🛡️ 管理權限身分組：{rRole.Mention}\n⏰ 紀錄時區：台北 (UTC+8)");
                        break;

                    case "addpoint":
                        // 權限檢查：檢查執行者是否有被設定的管理員身分組
                        if (botConfig == null) { await command.FollowupAsync("⚠️ 請先使用 `/setup-roca` 進行綁定設定。"); return; }
                        var executer = (SocketGuildUser)command.User;
                        if (!executer.Roles.Any(r => r.Id == botConfig.AdminRoleId) && executer.Id != guildChannel.Guild.OwnerId)
                        {
                            await command.FollowupAsync($"❌ 你沒有權限！只有持有 <@&{botConfig.AdminRoleId}> 身分組的人可以發放點數。");
                            return;
                        }

                        var targetUser = (SocketGuildUser)command.Data.Options.First(x => x.Name == "user").Value;
                        string targetName = targetUser.Nickname ?? targetUser.Username;
                        int pts = Convert.ToInt32((long)command.Data.Options.First(x => x.Name == "points").Value);
                        string reason = (string)command.Data.Options.First(x => x.Name == "reason").Value;

                        // Roblox 群組驗證
                        bool isInGroup = await VerifyUserInRobloxGroup(targetName, botConfig.RobloxGroupId);
                        if (!isInGroup) { await command.FollowupAsync($"❌ 玩家 **{targetName}** 不在綁定的 Roblox 群組 ({botConfig.RobloxGroupId}) 內！"); return; }

                        // 寫入點數 (鎖定伺服器 ID)
                        var userRecord = await db.UserPoints.FirstOrDefaultAsync(u => u.GuildId == currentGuildId && u.RobloxUsername.ToLower() == targetName.ToLower());
                        if (userRecord == null) { userRecord = new UserPoint { GuildId = currentGuildId, RobloxUsername = targetName, Points = 0 }; db.UserPoints.Add(userRecord); }
                        userRecord.Points += pts;

                        db.PointLogs.Add(new PointLog
                        {
                            GuildId = currentGuildId,
                            RobloxUsername = targetName,
                            AdminName = command.User.Username,
                            PointsAdded = pts,
                            Reason = reason,
                            Timestamp = Program.GetTaipeiTime() // 強制台北時間
                        });

                        await db.SaveChangesAsync();
                        await command.FollowupAsync(embed: new EmbedBuilder()
                            .WithTitle("✅ 點數發放成功 (台北時區)")
                            .WithColor(Color.Green)
                            .AddField("玩家", targetName, true)
                            .AddField("數量", $"+{pts}", true)
                            .AddField("總計", $"{userRecord.Points}", true)
                            .AddField("備註", reason)
                            .WithFooter($"操作時間: {Program.GetTaipeiTime():yyyy-MM-dd HH:mm:ss}")
                            .Build());
                        break;

                    case "points":
                        var user = (SocketGuildUser)command.Data.Options.First().Value;
                        string name = user.Nickname ?? user.Username;
                        var record = await db.UserPoints.FirstOrDefaultAsync(u => u.GuildId == currentGuildId && u.RobloxUsername.ToLower() == name.ToLower());
                        await command.FollowupAsync($"📊 玩家 **{name}** 在本伺服器的點數為：**{(record?.Points ?? 0)}** 點。");
                        break;

                    case "viewall":
                        var all = await db.UserPoints.Where(u => u.GuildId == currentGuildId).OrderByDescending(u => u.Points).ToListAsync();
                        if (all.Count == 0) { await command.FollowupAsync("📭 本伺服器尚無紀錄。"); break; }
                        StringBuilder sb = new StringBuilder("```markdown\n# 🏆 伺服器點數排行榜 (台北時區)\n=============================================\n");
                        foreach (var u in all) sb.AppendLine($"{u.RobloxUsername,-20} | {u.Points} 點");
                        sb.Append("```");
                        await command.FollowupAsync(sb.ToString());
                        break;
                }
            }
            catch (Exception ex) { await command.FollowupAsync($"⚠️ 錯誤: {ex.Message}"); }
        }

        private async Task HandleInteractionAsync(SocketInteraction interaction) { /* 按鈕邏輯維持之前提供的刪除機制 */ }

        private async Task<bool> VerifyUserInRobloxGroup(string username, string groupId)
        {
            try
            {
                var userReq = new { usernames = new[] { username }, excludeBannedUsers = true };
                var content = new StringContent(JsonSerializer.Serialize(userReq), Encoding.UTF8, "application/json");
                var userRes = await _http.PostAsync("https://users.roblox.com/v1/usernames/users", content);
                if (!userRes.IsSuccessStatusCode) return false;
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
    // 資料庫結構 (支援台北時區與多伺服器隔離)
    // ==========================================
    public class BotDbContext : DbContext
    {
        public DbSet<UserPoint> UserPoints { get; set; }
        public DbSet<PointLog> PointLogs { get; set; }
        public DbSet<BotConfig> Configs { get; set; }
        protected override void OnConfiguring(DbContextOptionsBuilder options) => options.UseSqlite("Data Source=rocapoints.db");
    }

    public class UserPoint { [Key] public int Id { get; set; } public ulong GuildId { get; set; } public string RobloxUsername { get; set; } public int Points { get; set; } }
    public class PointLog { [Key] public int Id { get; set; } public ulong GuildId { get; set; } public string RobloxUsername { get; set; } public string AdminName { get; set; } public int PointsAdded { get; set; } public string Reason { get; set; } public DateTime Timestamp { get; set; } public bool IsDeleted { get; set; } = false; }
    public class BotConfig { [Key] public ulong GuildId { get; set; } public string RobloxGroupId { get; set; } public ulong AdminRoleId { get; set; } }
}