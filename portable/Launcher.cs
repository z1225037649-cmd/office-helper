using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        bool snapshot = args.Length == 2 && args[0] == "--snapshot";
        bool smoke = args.Length == 2 && args[0] == "--smoke-wps";
        bool created;
        string sid = WindowsIdentity.GetCurrent().User.Value;
        using (var mutex = new Mutex(true, "Local\\OfficeHelperPortable-" + sid + (snapshot ? "-snapshot" : ""), out created))
        {
            if (!created) { MessageBox.Show("小帮手已经在运行，请从任务栏打开它。", "办公小帮手"); return; }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new HelperWindow(snapshot ? args[1] : null, smoke ? args[1] : null));
        }
    }
}

internal sealed class HelperWindow : Form
{
    private readonly string root = AppDomain.CurrentDomain.BaseDirectory;
    private readonly JavaScriptSerializer json = new JavaScriptSerializer();
    private readonly CheckBox wps = new CheckBox { Text = "WPS 文字 / 表格", AutoSize = true };
    private readonly CheckBox office = new CheckBox { Text = "Microsoft Word / Excel", AutoSize = true };
    private readonly Button start = new Button { Text = "开始使用", Height = 54, Dock = DockStyle.Fill };
    private readonly Button repair = new Button { Text = "检查与修复", Height = 44, Dock = DockStyle.Fill };
    private readonly Button stop = new Button { Text = "停止助手", Height = 44, Dock = DockStyle.Fill, Enabled = false };
    private readonly Button openWps = new Button { Text = "打开 WPS", Dock = DockStyle.Fill, Enabled = false };
    private readonly Button openWord = new Button { Text = "试用 Word", Dock = DockStyle.Fill, Enabled = false };
    private readonly Button openExcel = new Button { Text = "试用 Excel", Dock = DockStyle.Fill, Enabled = false };
    private readonly Label status = new Label { Text = "正在检查这台电脑…", AutoSize = true, MaximumSize = new Size(620, 0) };
    private readonly Label certStatus = new Label { Text = "选择常用的办公软件，点击开始。", AutoSize = true, MaximumSize = new Size(620, 0) };
    private readonly ProgressBar progress = new ProgressBar { Dock = DockStyle.Fill, Height = 8, Visible = false, Style = ProgressBarStyle.Marquee };
    private readonly LinkLabel remove = new LinkLabel { Text = "移除本机配置", AutoSize = true };
    private readonly System.Windows.Forms.Timer maintenance = new System.Windows.Forms.Timer { Interval = 6 * 60 * 60 * 1000 };
    private Dictionary<string, object> inspection;
    private Process server;
    private bool busy, closing, running, startedWps, startedOffice;
    private string lastServerError = "";

    internal HelperWindow(string snapshotPath, string smokePath)
    {
        Text = "办公小帮手 · 便携版";
        Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        Font = new Font("Microsoft YaHei UI", 11F);
        BackColor = Color.FromArgb(247, 249, 247);
        ForeColor = Color.FromArgb(35, 50, 42);
        StartPosition = FormStartPosition.CenterScreen;
        AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new Size(720, 720);
        MinimumSize = new Size(720, 720);
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(32, 25, 32, 22), ColumnCount = 1, RowCount = 11 };
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 62));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 61));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 16));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 68));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 28));
        layout.Controls.Add(new Label { Text = "办公轻松一点", Font = new Font(Font.FontFamily, 24, FontStyle.Bold), AutoSize = true }, 0, 0);
        layout.Controls.Add(new Label { Text = "启动后，在办公软件里使用你的“小帮手”。", AutoSize = true }, 0, 1);
        var choices = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(0, 12, 0, 0) };
        wps.Margin = new Padding(0, 0, 30, 0); choices.Controls.Add(wps); choices.Controls.Add(office);
        layout.Controls.Add(choices, 0, 2);
        start.BackColor = Color.FromArgb(35, 102, 83); start.ForeColor = Color.White; start.FlatStyle = FlatStyle.Flat; start.Font = new Font(Font.FontFamily, 14, FontStyle.Bold);
        layout.Controls.Add(start, 0, 3); layout.Controls.Add(progress, 0, 4);
        layout.Controls.Add(status, 0, 5); layout.Controls.Add(certStatus, 0, 6);
        var opens = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 1 };
        opens.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        for (int i = 0; i < 3; i++) opens.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.33F));
        opens.Controls.Add(openWps); opens.Controls.Add(openWord); opens.Controls.Add(openExcel); layout.Controls.Add(opens, 0, 7);
        var note = new Label { Text = "首次配置后，请保存文件并重新打开办公软件。\n保持本窗口运行；关闭后，插件会暂时断开。", AutoSize = true, ForeColor = Color.FromArgb(85, 100, 92), Margin = new Padding(0, 16, 0, 0) };
        layout.Controls.Add(note, 0, 8);
        var actions = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
        actions.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50)); actions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
        actions.Controls.Add(repair); actions.Controls.Add(stop); layout.Controls.Add(actions, 0, 9);
        layout.Controls.Add(remove, 0, 10); Controls.Add(layout);
        start.Click += async (s, e) => await StartHelper();
        repair.Click += async (s, e) => { SetBusy(true); try { if (running) await StopServer(); } finally { SetBusy(false); } await StartHelper(); };
        stop.Click += async (s, e) => { SetBusy(true); try { await StopServer(); status.Text = "助手已停止，文件仍在办公软件中正常打开。"; } finally { SetBusy(false); } };
        openWps.Click += async (s, e) => await OpenHost("OpenWps");
        openWord.Click += async (s, e) => await OpenHost("OpenWord");
        openExcel.Click += async (s, e) => await OpenHost("OpenExcel");
        remove.LinkClicked += async (s, e) => await RemoveConfiguration();
        maintenance.Tick += async (s, e) => {
            if (busy || !running || !startedOffice) return;
            bool renew = false; SetBusy(true);
            try {
                inspection = await Controller("Inspect", "Both", false);
                var cert = GetObject(inspection, "certificate");
                renew = Boolean(cert, "needsRenewal") || !Boolean(cert, "trusted");
                if (renew) await StopServer();
            } catch (Exception ex) { status.Text = "自动检查未完成，可以点击“检查与修复”。\n" + Explain(ex.Message); }
            finally { SetBusy(false); }
            if (renew && !closing) await StartHelper();
        };
        FormClosing += (s, e) => {
            if (busy && !closing) { e.Cancel = true; status.Text = "正在完成配置，请稍候再关闭。"; return; }
            closing = true; maintenance.Stop(); StopServer().GetAwaiter().GetResult();
        };
        Shown += async (s, e) => {
            if (snapshotPath != null) {
                wps.Checked = true; office.Checked = true; status.Text = "已找到 WPS 和 Microsoft Office。"; certStatus.Text = "首次启用 Office 时，会说明本机连接配置。";
                using (var bitmap = new Bitmap(Width, Height)) { DrawToBitmap(bitmap, new Rectangle(0, 0, Width, Height)); bitmap.Save(snapshotPath); }
                Close(); return;
            }
            SetBusy(true);
            try {
                inspection = await Controller("Inspect", "Both", false);
                var apps = GetObject(inspection, "applications");
                string preference = TextValue(inspection, "hosts");
                wps.Checked = preference != "Office" && !String.IsNullOrEmpty(TextValue(apps, "wps"));
                office.Checked = preference != "Wps" && (!String.IsNullOrEmpty(TextValue(apps, "word")) || !String.IsNullOrEmpty(TextValue(apps, "excel")));
                status.Text = wps.Checked || office.Checked ? "已找到办公软件，选择后点击“开始使用”。" : "没有找到 WPS 或 Microsoft Office，请先安装其中一种。";
                ShowCertificate();
            } catch (Exception ex) { status.Text = Explain(ex.Message); }
            finally { SetBusy(false); }
            if (smokePath != null) {
                wps.Checked = true; office.Checked = false;
                try {
                    await StartHelper();
                    if (!running) throw new Exception(status.Text);
                    string body;
                    using (var client = new System.Net.WebClient()) body = await client.DownloadStringTaskAsync("http://127.0.0.1:18765/health");
                    if (!body.Contains("office-helper-portable")) throw new Exception("Wrong local server");
                    SetBusy(true); await StopServer(); SetBusy(false);
                    File.WriteAllText(smokePath, json.Serialize(new { started = true, healthy = true, stopped = true }), Encoding.UTF8);
                } catch (Exception ex) { File.WriteAllText(smokePath, json.Serialize(new { error = ex.Message }), Encoding.UTF8); }
                Close();
            }
        };
    }

    private static Dictionary<string, object> GetObject(Dictionary<string, object> source, string name) { object value; return source != null && source.TryGetValue(name, out value) ? value as Dictionary<string, object> : null; }
    private static bool Boolean(Dictionary<string, object> source, string name) { object value; return source != null && source.TryGetValue(name, out value) && value is bool && (bool)value; }
    private static string TextValue(Dictionary<string, object> source, string name) { object value; return source != null && source.TryGetValue(name, out value) && value != null ? Convert.ToString(value) : ""; }
    private void SetBusy(bool value) { busy = value; progress.Visible = value; UpdateButtons(); }
    private void UpdateButtons() {
        start.Enabled = !busy && !running; repair.Enabled = !busy; stop.Enabled = !busy && running; remove.Enabled = !busy;
        wps.Enabled = !busy && !running; office.Enabled = !busy && !running;
        openWps.Enabled = !busy && running && startedWps;
        var apps = GetObject(inspection, "applications");
        openWord.Enabled = !busy && running && startedOffice && !String.IsNullOrEmpty(TextValue(apps, "word"));
        openExcel.Enabled = !busy && running && startedOffice && !String.IsNullOrEmpty(TextValue(apps, "excel"));
    }
    private void ShowCertificate() {
        var cert = GetObject(inspection, "certificate");
        DateTime expiry;
        certStatus.Text = Boolean(cert, "valid") && Boolean(cert, "trusted") && DateTime.TryParse(TextValue(cert, "expiresAt"), out expiry)
            ? "Office 连接有效至 " + expiry.ToLocalTime().ToString("yyyy-MM-dd") + "。续期时可能需要确认系统提示。"
            : "只使用 WPS 无需证书；Office 会在首次启用时完成配置。";
    }
    private async Task StartHelper() {
        if (busy || running) return;
        if (!wps.Checked && !office.Checked) { status.Text = "请至少选择一种办公软件。"; return; }
        SetBusy(true);
        try {
            inspection = await Controller("Inspect", "Both", false);
            bool allow = false;
            if (office.Checked && !Boolean(inspection, "allowCertificateMaintenance")) {
                var answer = MessageBox.Show(this, "为了连接 Word / Excel，小帮手需要为这台电脑创建本地连接证书，并加入当前 Windows 用户的信任列表。\n\n证书有效期 30 天；允许后，小帮手会在运行时自动检查和续期。Windows 可能要求你确认信任新证书。私钥保留在这台电脑，不放进便携包。\n\n你随时可以点击“移除本机配置”撤销。是否允许？", "首次启用 Office", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (answer != DialogResult.Yes) { status.Text = "尚未启用 Office。你也可以只勾选 WPS 使用。"; return; }
                allow = true;
            }
            status.Text = "正在准备插件和本机连接，请稍候…";
            string hosts = wps.Checked && office.Checked ? "Both" : office.Checked ? "Office" : "Wps";
            var prepared = await Controller("Prepare", hosts, allow);
            await StartServer(TextValue(prepared, "configPath"));
            if (server == null || server.HasExited) throw new Exception("SERVER_STOPPED");
            running = true; startedWps = wps.Checked; startedOffice = office.Checked;
            inspection = await Controller("Inspect", hosts, false);
            status.Text = "小帮手已启动。现在可以打开办公软件使用。";
            ShowCertificate(); maintenance.Start();
        } catch (Exception ex) { StopServer().GetAwaiter().GetResult(); status.Text = Explain(ex.Message); }
        finally { SetBusy(false); }
    }
    private async Task OpenHost(string action) { SetBusy(true); try { await Controller(action, "Both", false); } catch (Exception ex) { status.Text = Explain(ex.Message); } finally { SetBusy(false); } }
    private async Task RemoveConfiguration() {
        if (MessageBox.Show(this, "将停止助手，并移除便携版的插件入口和本机连接证书。你的文档、试用文件和其他插件会保留。", "移除本机配置", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK) return;
        SetBusy(true);
        try { await StopServer(); await Controller("Remove", "Both", false); inspection = await Controller("Inspect", "Both", false); ShowCertificate(); status.Text = "本机配置已移除。保存文件并重新打开办公软件后生效。"; }
        catch (Exception ex) { status.Text = Explain(ex.Message); }
        finally { SetBusy(false); }
    }
    private Task<Dictionary<string, object>> Controller(string action, string hosts, bool allow) {
        return Task.Run(() => {
            var info = new ProcessStartInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell\\v1.0\\powershell.exe"));
            info.Arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File \"" + Path.Combine(root, "app\\controller.ps1") + "\" -Action " + action + " -Hosts " + hosts + (allow ? " -AllowTrust" : "");
            info.UseShellExecute = false; info.CreateNoWindow = true; info.RedirectStandardOutput = true; info.RedirectStandardError = true; info.StandardOutputEncoding = Encoding.UTF8; info.StandardErrorEncoding = Encoding.UTF8; info.WorkingDirectory = root;
            using (var process = Process.Start(info)) {
                var output = process.StandardOutput.ReadToEndAsync(); var errors = process.StandardError.ReadToEndAsync();
                process.WaitForExit(); Task.WaitAll(output, errors);
                Dictionary<string, object> result;
                try { result = json.Deserialize<Dictionary<string, object>>(output.Result.Trim()); }
                catch { throw new Exception("CONFIG_FAILED: " + errors.Result.Trim()); }
                if (process.ExitCode != 0 || !Boolean(result, "ok")) throw new Exception(TextValue(result, "error"));
                return result;
            }
        });
    }
    private async Task StartServer(string config) {
        var ready = new TaskCompletionSource<bool>(); lastServerError = "";
        var info = new ProcessStartInfo(Path.Combine(root, "runtime\\node.exe"), "\"" + Path.Combine(root, "app\\server.mjs") + "\" \"" + config + "\"");
        info.UseShellExecute = false; info.CreateNoWindow = true; info.RedirectStandardInput = true; info.RedirectStandardOutput = true; info.RedirectStandardError = true; info.StandardOutputEncoding = Encoding.UTF8; info.StandardErrorEncoding = Encoding.UTF8; info.WorkingDirectory = root;
        server = new Process { StartInfo = info, EnableRaisingEvents = true };
        var launching = server;
        server.OutputDataReceived += (s, e) => { if (e.Data != null && e.Data.Contains("\"ready\":true")) ready.TrySetResult(true); };
        server.ErrorDataReceived += (s, e) => { if (e.Data != null) lastServerError = e.Data; };
        server.Exited += (s, e) => {
            ready.TrySetException(new Exception(lastServerError.Length > 0 ? lastServerError : "SERVER_STOPPED"));
            if (!closing && IsHandleCreated) BeginInvoke(new Action(() => { if (running && Object.ReferenceEquals(server, launching)) { running = false; status.Text = "连接服务已停止，请点击“开始使用”重试。"; UpdateButtons(); } }));
        };
        server.Start(); server.BeginOutputReadLine(); server.BeginErrorReadLine();
        if (await Task.WhenAny(ready.Task, Task.Delay(20000)) != ready.Task) throw new Exception("SERVER_TIMEOUT");
        await ready.Task;
    }
    private Task StopServer() {
        running = false; maintenance.Stop(); var owned = server; server = null;
        return Task.Run(() => {
            if (owned == null) return;
            try { if (!owned.HasExited) { owned.StandardInput.WriteLine("stop"); owned.StandardInput.Close(); if (!owned.WaitForExit(5000)) owned.Kill(); } } catch (InvalidOperationException) { } finally { owned.Dispose(); }
        });
    }
    private static string Explain(string value) {
        if (value.Contains("PORT_IN_USE") || value.Contains("EADDRINUSE")) return "连接端口正在被其他程序使用。请关闭另一份小帮手后重试；不会自动关闭其他程序。";
        if (value.Contains("TRUST_REQUIRED")) return "Office 需要先完成本机连接证书配置，请点击“检查与修复”。";
        if (value.Contains("OFFICE_NOT_FOUND")) return "没有找到 Microsoft Word / Excel，请取消勾选或先安装 Office。";
        if (value.Contains("WPS_NOT_FOUND")) return "没有找到 WPS，请取消勾选或先安装 WPS。";
        if (value.Contains("PACKAGE_INCOMPLETE")) return "便携包不完整，请完整解压后再启动，不要只复制启动程序。";
        if (value.Contains("REMOVE_PARTIAL")) return "部分配置未能移除，其他项目已尝试清理。请保留本程序并重试。\n" + value;
        if (value.Contains("WPS_CONFIG_UNEXPECTED")) return "WPS 插件配置无法读取。已停止修改，请先检查 WPS 是否正常运行。";
        return "这一步没有完成，可以点击“检查与修复”。\n" + (value.Length > 240 ? value.Substring(0, 240) : value);
    }
}
