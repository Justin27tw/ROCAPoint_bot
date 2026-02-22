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

namespace ROCAPointBot
{
    // ==========================================
    // 網頁伺服器進入點 (用來保持 24 小時清醒)
    // ==========================================
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            // 註冊我們的 Discord 機器人作為「背景服務」
            builder.Services.AddHostedService<DiscordBotService>();

            var app = builder.Build();

            // 建立網頁路由：讓 UptimeRobot 能夠每 5 分鐘來訪問這個網址，防止伺服器休眠
            app.MapGet("/", () => "✅ ROCA Discord Bot is currently online and running 24/7!");
            app.MapGet("/ping", () => "pong");

            app.Run();
        }
    }

    // ==========================================
    // Discord 機器人核心邏輯 (在背景默默執行)
    // ==========================================
    public class DiscordBotService : BackgroundService
    {
        private DiscordSocketClient _client;
        private static readonly HttpClient _http = new HttpClient();

        // 1. 先宣告一個變數來裝 Token (不要在這裡給值)
        private readonly string _discordToken;

        // 2. 加入這段「建構子」，讓系統自動去抓取安全設定檔裡的密碼
        public DiscordBotService(Microsoft.Extensions.Configuration.IConfiguration config)
        {
            // 系統會自動去尋找名叫 "DiscordToken" 的密碼
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
            _client.Log += Log;
            _client.Ready += RegisterCommandsAsync;
            _client.SlashCommandExecuted += HandleSlashCommandAsync;
            _client.InteractionCreated += HandleInteractionAsync;

            await _client.LoginAsync(TokenType.Bot, _discordToken);
            await _client.StartAsync();

            // 讓背景服務持續執行，直到伺服器關閉
            await Task.Delay(Timeout.Infinite, stoppingToken);
        }

        private Task Log(LogMessage arg)
        {
            Console.WriteLine(arg.ToString());
            return Task.CompletedTask;
        }

        private async Task RegisterCommandsAsync()
        {
            var commands = new List<ApplicationCommandProperties>
            {
                new SlashCommandBuilder().WithName("bindgroup").WithDescription("🔗 綁定此伺服器的 Roblox 群組").AddOption("group_id", ApplicationCommandOptionType.String, "Roblox 群組 ID", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("points").WithDescription("📊 查詢單一玩家點數").AddOption("user", ApplicationCommandOptionType.User, "選擇 Discord 使用者", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("setquick").WithDescription("⚡ 設定快捷加點代碼").AddOption("code", ApplicationCommandOptionType.String, "代碼 (例如 A, event)", isRequired: true).AddOption("points", ApplicationCommandOptionType.Integer, "點數數量", isRequired: true).AddOption("reason", ApplicationCommandOptionType.String, "發放原因", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("quicklist").WithDescription("📜 查看所有快捷選項清單").Build(),
                new SlashCommandBuilder().WithName("addpoint").WithDescription("➕ 手動發放點數給玩家").AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true).AddOption("points", ApplicationCommandOptionType.Integer, "點數數量", isRequired: true).AddOption("reason", ApplicationCommandOptionType.String, "發放原因", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("addquick").WithDescription("⚡ 使用快捷代碼發放點數給玩家").AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true).AddOption("code", ApplicationCommandOptionType.String, "快捷代碼", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("history").WithDescription("📜 查詢玩家近期點數異動紀錄").AddOption("user", ApplicationCommandOptionType.User, "選擇玩家", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("delrecord").WithDescription("🚨 刪除指定的單筆點數紀錄").AddOption("id", ApplicationCommandOptionType.Integer, "要刪除的紀錄 ID", isRequired: true).Build(),
                new SlashCommandBuilder().WithName("viewall").WithDescription("🏆 顯示伺服器內所有人點數排行榜總覽").Build()
            };

            try
            {
                await _client.BulkOverwriteGlobalApplicationCommandsAsync(commands.ToArray());
                Console.WriteLine("✅ 成功向 Discord 註冊所有斜線指令！");
            }
            catch (Exception ex) { Console.WriteLine($"❌ 註冊指令時發生錯誤: {ex.Message}"); }
        }

        private async Task HandleSlashCommandAsync(SocketSlashCommand command)
        {
            if (!(command.Channel is SocketGuildChannel guildChannel)) { await command.RespondAsync("⚠️ 這些指令只能在伺服器中使用！", ephemeral: true); return; }

            await command.DeferAsync();

            ulong currentGuildId = guildChannel.Guild.Id;
            using var db = new BotDbContext();

            var botConfig = await db.Configs.FirstOrDefaultAsync(c => c.GuildId == currentGuildId);
            if (botConfig == null) { botConfig = new BotConfig { GuildId = currentGuildId, RobloxGroupId = "", MaxPointsPerAdd = 100 }; db.Configs.Add(botConfig); await db.SaveChangesAsync(); }

            try
            {
                switch (command.Data.Name)
                {
                    case "bindgroup":
                        botConfig.RobloxGroupId = (string)command.Data.Options.First().Value;
                        await db.SaveChangesAsync();
                        await command.FollowupAsync($"🔗 成功綁定 Roblox 群組 ID：**{botConfig.RobloxGroupId}**");
                        break;
                    case "points":
                        var pUser = (SocketGuildUser)command.Data.Options.First().Value; string pName = pUser.Nickname ?? pUser.Username;
                        var pRecord = await db.UserPoints.FirstOrDefaultAsync(u => u.GuildId == currentGuildId && u.RobloxUsername.ToLower() == pName.ToLower());
                        await command.FollowupAsync($"📊 Roblox 玩家 **{pName}** 目前的總點數為：**{(pRecord?.Points ?? 0)}** 點。");
                        break;
                    case "setquick":
                        string sCode = ((string)command.Data.Options.First(x => x.Name == "code").Value).ToLower();
                        int sPts = Convert.ToInt32((long)command.Data.Options.First(x => x.Name == "points").Value);
                        string sReason = (string)command.Data.Options.First(x => x.Name == "reason").Value;
                        var existing = await db.QuickTemplates.FirstOrDefaultAsync(t => t.GuildId == currentGuildId && t.Code == sCode);
                        if (existing != null) { existing.Points = sPts; existing.Reason = sReason; } else db.QuickTemplates.Add(new QuickTemplate { GuildId = currentGuildId, Code = sCode, Points = sPts, Reason = sReason });
                        await db.SaveChangesAsync(); await command.FollowupAsync($"✅ 快捷選項設定成功！\n日後只需使用 `/addquick` 並輸入代碼 `{sCode}` 即可快速發放 **{sPts} 點** (原因: {sReason})");
                        break;
                    case "quicklist":
                        var templates = await db.QuickTemplates.Where(t => t.GuildId == currentGuildId).ToListAsync();
                        if (templates.Count == 0) { await command.FollowupAsync("📭 目前還沒有設定任何快捷選項。"); break; }
                        var qEmbed = new EmbedBuilder().WithTitle("⚡ 快捷發放選項總覽").WithColor(Color.Blue);
                        foreach (var t in templates) qEmbed.AddField($"代碼: `{t.Code}`", $"➕ {t.Points} 點 | 📝 {t.Reason}");
                        await command.FollowupAsync(embed: qEmbed.Build());
                        break;
                    case "addpoint":
                    case "addquick":
                        if (string.IsNullOrEmpty(botConfig.RobloxGroupId)) { await command.FollowupAsync("⚠️ 尚未綁定群組！請先使用 `/bindgroup`。"); break; }
                        var targetUser = (SocketGuildUser)command.Data.Options.First(x => x.Name == "user").Value;
                        string targetRobloxUser = targetUser.Nickname ?? targetUser.Username;
                        int pointsToAdd = 0; string reason = "";

                        if (command.Data.Name == "addquick")
                        {
                            string code = ((string)command.Data.Options.First(x => x.Name == "code").Value).ToLower();
                            var tmpl = await db.QuickTemplates.FirstOrDefaultAsync(t => t.GuildId == currentGuildId && t.Code == code);
                            if (tmpl == null) { await command.FollowupAsync($"⚠️ 找不到快捷代碼 `{code}`！"); break; }
                            pointsToAdd = tmpl.Points; reason = tmpl.Reason + " (快捷)";
                        }
                        else
                        {
                            pointsToAdd = Convert.ToInt32((long)command.Data.Options.First(x => x.Name == "points").Value);
                            reason = (string)command.Data.Options.First(x => x.Name == "reason").Value;
                        }

                        if (pointsToAdd > botConfig.MaxPointsPerAdd) { await command.FollowupAsync($"❌ 違反點數規範！此伺服器單次最高發放上限為 **{botConfig.MaxPointsPerAdd}** 點。"); break; }
                        bool isInGroup = await VerifyUserInRobloxGroup(targetRobloxUser, botConfig.RobloxGroupId);
                        if (!isInGroup) { await command.FollowupAsync($"❌ 發放失敗：玩家 **{targetRobloxUser}** 不存在或不在群組內！\n*(自動讀取的名稱：{targetRobloxUser}，如果不對請檢查使用者的伺服器暱稱)*"); break; }

                        var userRecord = await db.UserPoints.FirstOrDefaultAsync(u => u.GuildId == currentGuildId && u.RobloxUsername.ToLower() == targetRobloxUser.ToLower());
                        if (userRecord == null) { userRecord = new UserPoint { GuildId = currentGuildId, RobloxUsername = targetRobloxUser, Points = 0, DiscordId = targetUser.Id }; db.UserPoints.Add(userRecord); }
                        userRecord.Points += pointsToAdd;
                        db.PointLogs.Add(new PointLog { GuildId = currentGuildId, TargetRobloxUsername = targetRobloxUser, AdminDiscordName = command.User.Username, PointsAdded = pointsToAdd, Reason = reason, Timestamp = DateTime.UtcNow.AddHours(8) });
                        await db.SaveChangesAsync();
                        var embed = new EmbedBuilder().WithTitle("✅ 點數發放成功").WithColor(Color.Green).AddField("👤 玩家名稱", targetRobloxUser, true).AddField("🔢 本次發放", $"+{pointsToAdd} 點", true).AddField("📊 目前總計", $"{userRecord.Points} 點", true).AddField("📝 發放原因", reason);
                        await command.FollowupAsync(embed: embed.Build());
                        break;
                    case "history":
                        var hUser = (SocketGuildUser)command.Data.Options.First().Value; string hName = hUser.Nickname ?? hUser.Username;
                        var logs = await db.PointLogs.Where(l => l.GuildId == currentGuildId && l.TargetRobloxUsername.ToLower() == hName.ToLower() && !l.IsDeleted).OrderByDescending(l => l.Timestamp).Take(10).ToListAsync();
                        if (logs.Count == 0) { await command.FollowupAsync($"📭 玩家 **{hName}** 目前沒有近期紀錄。"); break; }
                        var hEmbed = new EmbedBuilder().WithTitle($"📜 {hName} 的近期點數紀錄").WithColor(Color.LightOrange);
                        foreach (var log in logs) hEmbed.AddField($"紀錄 ID: `#{log.Id}` | 時間: {log.Timestamp:MM/dd HH:mm}", $"➕ 增加 **{log.PointsAdded}** 點\n📝 {log.Reason}\n👨‍💼 由 {log.AdminDiscordName} 發放");
                        await command.FollowupAsync(embed: hEmbed.Build());
                        break;
                    case "delrecord":
                        int delId = Convert.ToInt32((long)command.Data.Options.First().Value);
                        var delLog = await db.PointLogs.FirstOrDefaultAsync(l => l.Id == delId && l.GuildId == currentGuildId);
                        if (delLog == null) { await command.FollowupAsync("⚠️ 找不到這筆紀錄 ID。"); break; }
                        if (delLog.IsDeleted) { await command.FollowupAsync("⚠️ 這筆紀錄已經被刪除過了！"); break; }
                        var builder = new ComponentBuilder().WithButton("⚠️ 確認刪除", $"del_confirm_{delLog.Id}", ButtonStyle.Danger).WithButton("取消", $"del_cancel_{delLog.Id}", ButtonStyle.Secondary);
                        await command.FollowupAsync($"🚨 **【刪除確認】** 您即將刪除紀錄 **#{delLog.Id}**\n玩家：{delLog.TargetRobloxUsername} | 點數：+{delLog.PointsAdded}\n*確認後點數將自動扣回。*", components: builder.Build());
                        break;
                    case "viewall":
                        var allUsers = await db.UserPoints.Where(u => u.GuildId == currentGuildId).OrderByDescending(u => u.Points).ToListAsync();
                        if (allUsers.Count == 0) { await command.FollowupAsync("📭 目前此伺服器內還沒有任何人有紀錄。"); break; }
                        StringBuilder sb = new StringBuilder(); sb.AppendLine("```markdown\n# 🏆 Roblox 群組點數總覽 (共紀錄 " + allUsers.Count + " 人)\n=============================================");
                        int rank = 1;
                        foreach (var u in allUsers)
                        {
                            string line = $"{rank,-3}. {u.RobloxUsername,-20} | {u.Points} 點\n";
                            if (sb.Length + line.Length > 1900) { sb.AppendLine("```"); await command.FollowupAsync(sb.ToString()); sb.Clear(); sb.AppendLine("```markdown"); }
                            sb.Append(line); rank++;
                        }
                        if (sb.Length > 12) { sb.AppendLine("```"); await command.FollowupAsync(sb.ToString()); }
                        break;
                }
            }
            catch (Exception ex) { await command.FollowupAsync($"⚠️ 發生錯誤：{ex.Message}"); }
        }

        private async Task HandleInteractionAsync(SocketInteraction interaction)
        {
            if (interaction is SocketMessageComponent component)
            {
                string customId = component.Data.CustomId;
                if (customId.StartsWith("del_cancel_")) { await component.UpdateAsync(msg => { msg.Content = "❌ 刪除動作已取消。"; msg.Components = new ComponentBuilder().Build(); }); }
                else if (customId.StartsWith("del_confirm_"))
                {
                    int logId = int.Parse(customId.Replace("del_confirm_", ""));
                    using var db = new BotDbContext(); var log = await db.PointLogs.FirstOrDefaultAsync(l => l.Id == logId);
                    if (log == null || log.IsDeleted) { await component.UpdateAsync(msg => { msg.Content = "⚠️ 紀錄已不存在。"; msg.Components = null; }); return; }
                    log.IsDeleted = true; log.DeletedBy = component.User.Username; log.DeletedAt = DateTime.UtcNow.AddHours(8);
                    var userRecord = await db.UserPoints.FirstOrDefaultAsync(u => u.GuildId == log.GuildId && u.RobloxUsername == log.TargetRobloxUsername);
                    if (userRecord != null) { userRecord.Points = Math.Max(0, userRecord.Points - log.PointsAdded); }
                    await db.SaveChangesAsync();
                    await component.UpdateAsync(msg => { msg.Content = $"✅ **已成功刪除紀錄 #{logId}**。\n已將玩家 **{log.TargetRobloxUsername}** 扣回 {log.PointsAdded} 點。"; msg.Components = new ComponentBuilder().Build(); });
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
                if (!userRes.IsSuccessStatusCode) return false;
                var userJson = await userRes.Content.ReadAsStringAsync(); using var userDoc = JsonDocument.Parse(userJson);
                var dataArray = userDoc.RootElement.GetProperty("data"); if (dataArray.GetArrayLength() == 0) return false;
                long userId = dataArray[0].GetProperty("id").GetInt64();
                var groupRes = await _http.GetAsync($"https://groups.roblox.com/v1/users/{userId}/groups/roles");
                if (!groupRes.IsSuccessStatusCode) return false;
                var groupJson = await groupRes.Content.ReadAsStringAsync(); using var groupDoc = JsonDocument.Parse(groupJson);
                var groupArray = groupDoc.RootElement.GetProperty("data");
                foreach (var group in groupArray.EnumerateArray()) { if (group.GetProperty("group").GetProperty("id").GetInt64().ToString() == groupId) return true; }
                return false;
            }
            catch { return false; }
        }
    }

    // ==========================================
    // 資料庫結構定義
    // ==========================================
    public class BotDbContext : DbContext
    {
        public DbSet<UserPoint> UserPoints { get; set; }
        public DbSet<PointLog> PointLogs { get; set; }
        public DbSet<BotConfig> Configs { get; set; }
        public DbSet<QuickTemplate> QuickTemplates { get; set; }
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder) { optionsBuilder.UseSqlite("Data Source=rocapoints.db"); }
    }
    public class QuickTemplate { [Key] public int Id { get; set; } public ulong GuildId { get; set; } public string Code { get; set; } public int Points { get; set; } public string Reason { get; set; } }
    public class UserPoint { [Key] public int Id { get; set; } public ulong GuildId { get; set; } public string RobloxUsername { get; set; } public ulong? DiscordId { get; set; } public int Points { get; set; } }
    public class PointLog { [Key] public int Id { get; set; } public ulong GuildId { get; set; } public string TargetRobloxUsername { get; set; } public string AdminDiscordName { get; set; } public int PointsAdded { get; set; } public string Reason { get; set; } public DateTime Timestamp { get; set; } public bool IsDeleted { get; set; } = false; public string DeletedBy { get; set; } = null; public DateTime? DeletedAt { get; set; } = null; }
    public class BotConfig { [Key] public ulong GuildId { get; set; } public string RobloxGroupId { get; set; } public int MaxPointsPerAdd { get; set; } }
}