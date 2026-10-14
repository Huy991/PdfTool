// Bộ cài đặt / gỡ cài đặt PdfTool.
// - Bản có nhúng payload.zip (PdfTool-Setup.exe): cài đặt cho người dùng hiện tại, không cần quyền Admin.
// - Bản không có payload (uninstall.exe, nằm trong thư mục cài): gỡ cài đặt.
// Biên dịch: chạy ..\installer.bat

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace PdfToolSetup
{
    static class Common
    {
        public const string AppName = "PdfTool";
        public const string DisplayName = "PdfTool - Bộ công cụ PDF";
        public const string Version = "1.0.0";
        public const string Publisher = "Ngo Quang Huy";
        public const string UninstallKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\PdfTool";
        public const string FileList = "install-files.txt";

        public static string DesktopLink { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), AppName + ".lnk"); } }
        public static string StartMenuLink { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), AppName + ".lnk"); } }

        static object Com(object o, string name, BindingFlags f, params object[] args)
        {
            return o.GetType().InvokeMember(name, f, null, o, args);
        }

        public static void CreateShortcut(string lnk, string target, string desc)
        {
            object sh = Activator.CreateInstance(Type.GetTypeFromProgID("WScript.Shell"));
            object sc = Com(sh, "CreateShortcut", BindingFlags.InvokeMethod, lnk);
            Com(sc, "TargetPath", BindingFlags.SetProperty, target);
            Com(sc, "WorkingDirectory", BindingFlags.SetProperty, Path.GetDirectoryName(target));
            Com(sc, "Description", BindingFlags.SetProperty, desc);
            Com(sc, "IconLocation", BindingFlags.SetProperty, target + ",0");
            Com(sc, "Save", BindingFlags.InvokeMethod);
            Marshal.FinalReleaseComObject(sc);
            Marshal.FinalReleaseComObject(sh);
        }

        public static string ShortcutTarget(string lnk)
        {
            try
            {
                if (!File.Exists(lnk)) return null;
                object sh = Activator.CreateInstance(Type.GetTypeFromProgID("WScript.Shell"));
                object sc = Com(sh, "CreateShortcut", BindingFlags.InvokeMethod, lnk);
                string t = (string)Com(sc, "TargetPath", BindingFlags.GetProperty);
                Marshal.FinalReleaseComObject(sc);
                Marshal.FinalReleaseComObject(sh);
                return t;
            }
            catch { return null; }
        }

        // Các tiến trình PdfTool đang chạy từ thư mục cài; trả về false nếu người dùng không đóng
        public static bool CloseRunning(string dir, IWin32Window owner)
        {
            string exe = Path.Combine(dir, AppName + ".exe");
            Func<List<Process>> find = () => Process.GetProcessesByName(AppName).Where(p =>
            {
                try { return string.Equals(p.MainModule.FileName, exe, StringComparison.OrdinalIgnoreCase); } catch { return false; }
            }).ToList();
            var running = find();
            if (running.Count == 0) return true;
            if (MessageBox.Show(owner, "PdfTool đang mở. Đóng PdfTool để tiếp tục?", DisplayName, MessageBoxButtons.OKCancel, MessageBoxIcon.Information) != DialogResult.OK)
                return false;
            foreach (var p in running) { try { p.CloseMainWindow(); p.WaitForExit(10000); } catch { } }
            if (find().Count == 0) return true;
            MessageBox.Show(owner, "Không đóng được PdfTool (có thể đang hỏi lưu tài liệu). Hãy tự đóng rồi thử lại.", DisplayName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        public static Icon AppIcon()
        {
            try { return Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { return null; }
        }
    }

    // ------------------------------------------------------------------
    // Cài đặt
    // ------------------------------------------------------------------
    class SetupForm : Form
    {
        readonly float scale;
        int S(int v) { return (int)(v * scale); }
        readonly TextBox txtDir;
        readonly CheckBox chkDesktop, chkLaunch;
        readonly ProgressBar progress;
        readonly Label lblStatus;
        readonly Button btnInstall, btnCancel, btnBrowse;
        bool done;

        public SetupForm()
        {
            using (var g = CreateGraphics()) scale = g.DpiX / 96f;
            Text = "Cài đặt " + Common.DisplayName;
            Font = new Font("Segoe UI", 9f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(S(540), S(330));
            Icon = Common.AppIcon();
            BackColor = Color.White;

            var banner = new Panel { Dock = DockStyle.Top, Height = S(84), BackColor = Color.FromArgb(0, 102, 190) };
            var logo = new PictureBox { Location = new Point(S(18), S(14)), Size = new Size(S(56), S(56)), SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.Transparent };
            if (Icon != null) logo.Image = new Icon(Icon, S(64), S(64)).ToBitmap();
            banner.Controls.Add(logo);
            banner.Controls.Add(new Label { Text = "PdfTool " + Common.Version, ForeColor = Color.White, Font = new Font("Segoe UI Semibold", 15f), AutoSize = true, Location = new Point(S(86), S(14)), BackColor = Color.Transparent });
            banner.Controls.Add(new Label { Text = "Bộ công cụ PDF: soạn, gộp, tách, nén, OCR, đánh số trang, mật khẩu...", ForeColor = Color.FromArgb(220, 235, 250), AutoSize = true, Location = new Point(S(88), S(48)), BackColor = Color.Transparent });

            int y = S(100);
            Controls.Add(new Label { Text = "Thư mục cài đặt:", AutoSize = true, Location = new Point(S(20), y) });
            txtDir = new TextBox
            {
                Location = new Point(S(20), y + S(22)), Width = S(450),
                Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", Common.AppName)
            };
            btnBrowse = new Button { Text = "...", Location = new Point(S(476), y + S(20)), Size = new Size(S(44), S(26)) };
            btnBrowse.Click += (s, e) =>
            {
                using (var d = new FolderBrowserDialog { Description = "Chọn thư mục cài đặt (sẽ tạo thư mục con PdfTool)" })
                    if (d.ShowDialog(this) == DialogResult.OK) txtDir.Text = Path.Combine(d.SelectedPath, Common.AppName);
            };
            chkDesktop = new CheckBox { Text = "Tạo biểu tượng trên màn hình (Desktop)", Checked = true, AutoSize = true, Location = new Point(S(20), y + S(60)) };
            chkLaunch = new CheckBox { Text = "Mở PdfTool sau khi cài xong", Checked = true, AutoSize = true, Location = new Point(S(20), y + S(84)) };
            progress = new ProgressBar { Location = new Point(S(20), y + S(120)), Width = S(500), Height = S(18) };
            lblStatus = new Label { AutoSize = true, Location = new Point(S(20), y + S(142)), ForeColor = Color.DimGray, Text = "Không cần quyền quản trị (Admin). Dung lượng khoảng 75 MB." };

            btnInstall = new Button { Text = "Cài đặt", Size = new Size(S(100), S(32)), Location = new Point(S(314), S(284)), BackColor = Color.FromArgb(0, 120, 212), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btnCancel = new Button { Text = "Hủy", Size = new Size(S(100), S(32)), Location = new Point(S(420), S(284)) };
            btnInstall.Click += (s, e) => { if (done) { Finish(); } else Install(); };
            btnCancel.Click += (s, e) => Close();
            AcceptButton = btnInstall;

            Controls.AddRange(new Control[] { txtDir, btnBrowse, chkDesktop, chkLaunch, progress, lblStatus, btnInstall, btnCancel });
            Controls.Add(banner);
        }

        void Install()
        {
            string dir = txtDir.Text.Trim().Trim('"');
            if (dir.Length == 0 || !Path.IsPathRooted(dir)) { MessageBox.Show(this, "Thư mục cài đặt không hợp lệ.", Text); return; }
            dir = Path.GetFullPath(dir);
            if (!Common.CloseRunning(dir, this)) return;

            btnInstall.Enabled = btnCancel.Enabled = btnBrowse.Enabled = txtDir.Enabled = chkDesktop.Enabled = false;
            bool desktop = chkDesktop.Checked;
            lblStatus.Text = "Đang cài đặt...";
            var t = new Thread(() =>
            {
                string error = null;
                try { DoInstall(dir, desktop); }
                catch (Exception ex) { error = ex.Message; }
                BeginInvoke(new Action(() =>
                {
                    if (error != null)
                    {
                        lblStatus.Text = "Cài đặt thất bại.";
                        MessageBox.Show(this, "Không cài được:\r\n" + error, Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        btnInstall.Enabled = btnCancel.Enabled = btnBrowse.Enabled = txtDir.Enabled = chkDesktop.Enabled = true;
                        return;
                    }
                    done = true;
                    installedDir = dir;
                    progress.Value = progress.Maximum;
                    lblStatus.Text = "Đã cài xong vào " + dir;
                    btnInstall.Text = "Hoàn tất";
                    btnInstall.Enabled = true;
                    btnCancel.Visible = false;
                }));
            });
            t.IsBackground = true;
            t.Start();
        }

        string installedDir;

        void Report(int value, int max)
        {
            BeginInvoke(new Action(() => { progress.Maximum = Math.Max(1, max); progress.Value = Math.Min(value, progress.Maximum); }));
        }

        void DoInstall(string dir, bool desktop)
        {
            Directory.CreateDirectory(dir);
            var installed = new List<string>();
            long total = 0;
            var asm = Assembly.GetExecutingAssembly();
            using (var s = asm.GetManifestResourceStream("payload.zip"))
            using (var zip = new ZipArchive(s, ZipArchiveMode.Read))
            {
                var entries = zip.Entries.Where(e => e.Name.Length > 0).ToList();
                int i = 0;
                foreach (var e in entries)
                {
                    // Bỏ thư mục gốc "PdfTool/" trong file zip
                    string rel = e.FullName.Replace('/', '\\');
                    int slash = rel.IndexOf('\\');
                    if (slash >= 0 && rel.Substring(0, slash).Equals(Common.AppName, StringComparison.OrdinalIgnoreCase)) rel = rel.Substring(slash + 1);
                    string dst = Path.GetFullPath(Path.Combine(dir, rel));
                    if (!dst.StartsWith(dir.TrimEnd('\\') + "\\", StringComparison.OrdinalIgnoreCase)) continue;
                    Directory.CreateDirectory(Path.GetDirectoryName(dst));
                    e.ExtractToFile(dst, true);
                    installed.Add(rel);
                    total += e.Length;
                    Report(++i, entries.Count + 2);
                }
            }

            string uninst = Path.Combine(dir, "uninstall.exe");
            using (var s = asm.GetManifestResourceStream("uninstall.exe"))
            using (var f = File.Create(uninst)) s.CopyTo(f);
            installed.Add("uninstall.exe");
            installed.Add(Common.FileList);
            File.WriteAllLines(Path.Combine(dir, Common.FileList), installed, Encoding.UTF8);

            string exe = Path.Combine(dir, Common.AppName + ".exe");
            Common.CreateShortcut(Common.StartMenuLink, exe, Common.DisplayName);
            if (desktop) Common.CreateShortcut(Common.DesktopLink, exe, Common.DisplayName);

            using (var k = Registry.CurrentUser.CreateSubKey(Common.UninstallKey))
            {
                k.SetValue("DisplayName", Common.DisplayName);
                k.SetValue("DisplayVersion", Common.Version);
                k.SetValue("Publisher", Common.Publisher);
                k.SetValue("DisplayIcon", exe);
                k.SetValue("InstallLocation", dir);
                k.SetValue("UninstallString", "\"" + uninst + "\"");
                k.SetValue("EstimatedSize", (int)(total / 1024), RegistryValueKind.DWord);
                k.SetValue("NoModify", 1, RegistryValueKind.DWord);
                k.SetValue("NoRepair", 1, RegistryValueKind.DWord);
            }
            Report(1, 1);
        }

        void Finish()
        {
            if (chkLaunch.Checked && installedDir != null)
                try { Process.Start(new ProcessStartInfo(Path.Combine(installedDir, Common.AppName + ".exe")) { WorkingDirectory = installedDir }); } catch { }
            Close();
        }
    }

    // ------------------------------------------------------------------
    // Gỡ cài đặt
    // ------------------------------------------------------------------
    static class Uninstaller
    {
        public static void Run(string[] args)
        {
            string self = Application.ExecutablePath;
            if (args.Length >= 2 && args[0] == "/run")
            {
                Remove(args[1]);
                // Tự xóa bản tạm của uninstall.exe sau khi thoát
                Process.Start(new ProcessStartInfo("cmd.exe", "/c ping -n 3 127.0.0.1 >nul & del \"" + self + "\"") { CreateNoWindow = true, UseShellExecute = false });
                return;
            }

            string dir = Path.GetDirectoryName(self);
            if (MessageBox.Show("Gỡ cài đặt " + Common.DisplayName + " khỏi máy?", Common.DisplayName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            if (!Common.CloseRunning(dir, null)) return;
            // Chạy bản sao trong thư mục tạm để có thể xóa được chính uninstall.exe
            string tmp = Path.Combine(Path.GetTempPath(), "pdftool_uninstall_" + Guid.NewGuid().ToString("N") + ".exe");
            File.Copy(self, tmp);
            Process.Start(new ProcessStartInfo(tmp, "/run \"" + dir + "\"") { UseShellExecute = false });
        }

        static void Remove(string dir)
        {
            Thread.Sleep(1000); // chờ uninstall.exe gốc thoát
            foreach (string lnk in new[] { Common.DesktopLink, Common.StartMenuLink })
            {
                string target = Common.ShortcutTarget(lnk);
                if (target != null && target.StartsWith(dir, StringComparison.OrdinalIgnoreCase))
                    try { File.Delete(lnk); } catch { }
            }
            try { Registry.CurrentUser.DeleteSubKeyTree(Common.UninstallKey, false); } catch { }

            // Chỉ xóa các file do bộ cài tạo ra, rồi xóa thư mục rỗng
            string list = Path.Combine(dir, Common.FileList);
            if (File.Exists(list))
            {
                foreach (string rel in File.ReadAllLines(list, Encoding.UTF8))
                {
                    if (rel.Length == 0 || rel.Contains("..")) continue;
                    try { File.Delete(Path.Combine(dir, rel)); } catch { }
                }
                foreach (string d in Directory.GetDirectories(dir, "*", SearchOption.AllDirectories).OrderByDescending(x => x.Length))
                    try { if (!Directory.EnumerateFileSystemEntries(d).Any()) Directory.Delete(d); } catch { }
                try { if (!Directory.EnumerateFileSystemEntries(dir).Any()) Directory.Delete(dir); } catch { }
            }
            MessageBox.Show("Đã gỡ cài đặt " + Common.DisplayName + ".", Common.DisplayName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    static class Program
    {
        [DllImport("user32.dll")] static extern bool SetProcessDPIAware();

        [STAThread]
        static void Main(string[] args)
        {
            try { SetProcessDPIAware(); } catch { }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            bool hasPayload = Assembly.GetExecutingAssembly().GetManifestResourceNames().Contains("payload.zip");
            try
            {
                if (hasPayload) Application.Run(new SetupForm());
                else Uninstaller.Run(args);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Common.DisplayName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
