// Giao diện PdfTool: bảng màu sáng/tối kiểu Windows 11 và các điều khiển tự vẽ
// (nút, ô nhập, ô chọn, ô đánh dấu, thẻ...). Màu được đọc lúc vẽ nên đổi giao diện không cần mở lại ứng dụng.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

namespace PdfTool
{
    enum ThemeMode { System, Light, Dark }
    enum Role { Window, Surface, Canvas, Input, Divider }
    enum TextRole { Muted, Accent }

    // Điều khiển tự vẽ: được gọi lại khi đổi giao diện
    interface IThemed { void ApplyTheme(); }

    class Palette
    {
        public bool Dark;
        public Color Window, Surface, Canvas, Border, Divider, Text, TextMuted, TextDisabled,
                     Accent, AccentHover, AccentPressed, OnAccent, AccentText, Hover, Pressed, Selection,
                     InputBg, InputBorder, InputBottom, ButtonBg, ButtonHover, ButtonPressed, ButtonBorder,
                     Menu, Warning, WarningBg, DonateBg, DonateBorder, DonateTitle;

        static Color C(int rgb) { return Color.FromArgb(255, (rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255); }

        public static Palette LightP()
        {
            return new Palette
            {
                Dark = false,
                Window = C(0xF3F3F3), Surface = C(0xFFFFFF), Canvas = C(0xE6E8EB), Border = C(0xE1E1E1), Divider = C(0xEAEAEA),
                Text = C(0x1B1B1B), TextMuted = C(0x5F5F5F), TextDisabled = C(0xA0A0A0),
                Accent = C(0x0067C0), AccentHover = C(0x1975C5), AccentPressed = C(0x3183CA), OnAccent = C(0xFFFFFF), AccentText = C(0x005FB8),
                Hover = Color.FromArgb(14, 0, 0, 0), Pressed = Color.FromArgb(24, 0, 0, 0), Selection = Color.FromArgb(34, 0x00, 0x67, 0xC0),
                InputBg = C(0xFFFFFF), InputBorder = C(0xD9D9D9), InputBottom = C(0x868686),
                ButtonBg = C(0xFDFDFD), ButtonHover = C(0xF3F3F3), ButtonPressed = C(0xEBEBEB), ButtonBorder = C(0xD3D3D3),
                Menu = C(0xF9F9F9), Warning = C(0x9D5D00), WarningBg = C(0xFFF4CE),
                DonateBg = C(0xFFF8EC), DonateBorder = C(0xF0DAB0), DonateTitle = C(0x8A5300)
            };
        }

        public static Palette DarkP()
        {
            return new Palette
            {
                Dark = true,
                Window = C(0x202020), Surface = C(0x2B2B2B), Canvas = C(0x181818), Border = C(0x3A3A3A), Divider = C(0x353535),
                Text = C(0xF3F3F3), TextMuted = C(0xAEAEAE), TextDisabled = C(0x6F6F6F),
                Accent = C(0x4CC2FF), AccentHover = C(0x47B1E8), AccentPressed = C(0x42A1D2), OnAccent = C(0x000000), AccentText = C(0x60CDFF),
                Hover = Color.FromArgb(20, 255, 255, 255), Pressed = Color.FromArgb(12, 255, 255, 255), Selection = Color.FromArgb(52, 0x4C, 0xC2, 0xFF),
                InputBg = C(0x323232), InputBorder = C(0x404040), InputBottom = C(0x9A9A9A),
                ButtonBg = C(0x373737), ButtonHover = C(0x3E3E3E), ButtonPressed = C(0x323232), ButtonBorder = C(0x464646),
                Menu = C(0x2C2C2C), Warning = C(0xFCE100), WarningBg = C(0x433519),
                DonateBg = C(0x342A1D), DonateBorder = C(0x5A4626), DonateTitle = C(0xF5C67E)
            };
        }
    }

    static class Theme
    {
        public static ThemeMode Mode = ThemeMode.System;
        public static Palette P = Palette.LightP();
        public static float Scale = 1f;
        public static string IconFamily = "Segoe MDL2 Assets";
        public static event Action Changed;

        static readonly ConditionalWeakTable<Control, object> bgRole = new ConditionalWeakTable<Control, object>();
        static readonly ConditionalWeakTable<Control, object> fgRole = new ConditionalWeakTable<Control, object>();
        static readonly ConditionalWeakTable<Control, object> hooked = new ConditionalWeakTable<Control, object>();
        static readonly Dictionary<string, Font> fonts = new Dictionary<string, Font>();

        public static int S(int v) { return (int)Math.Round(v * Scale); }
        public static float Sf(float v) { return v * Scale; }

        public static void Init()
        {
            using (var g = Graphics.FromHwnd(IntPtr.Zero)) Scale = g.DpiX / 96f;
            IconFamily = HasFont("Segoe Fluent Icons") ? "Segoe Fluent Icons" : "Segoe MDL2 Assets";
            string m = Settings.Get("theme", "system");
            Mode = m == "light" ? ThemeMode.Light : m == "dark" ? ThemeMode.Dark : ThemeMode.System;
            P = IsDark ? Palette.DarkP() : Palette.LightP();
            AppMode();
            SystemEvents.UserPreferenceChanged += (s, e) => { if (Mode == ThemeMode.System) Refresh(); };
        }

        static bool HasFont(string name)
        {
            try { using (var f = new Font(name, 10f)) return f.Name == name; }
            catch { return false; }
        }

        public static bool SystemDark
        {
            get
            {
                try
                {
                    using (var k = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                    {
                        object v = k != null ? k.GetValue("AppsUseLightTheme") : null;
                        return v is int && (int)v == 0;
                    }
                }
                catch { return false; }
            }
        }

        public static bool IsDark { get { return Mode == ThemeMode.Dark || (Mode == ThemeMode.System && SystemDark); } }

        public static void SetMode(ThemeMode m)
        {
            Mode = m;
            Settings.Set("theme", m == ThemeMode.Light ? "light" : m == ThemeMode.Dark ? "dark" : "system");
            Refresh();
        }

        static void Refresh()
        {
            bool dark = IsDark;
            if (dark == P.Dark) return;
            P = dark ? Palette.DarkP() : Palette.LightP();
            AppMode();
            var h = Changed;
            if (h != null) h();
        }

        // ---------------- phông chữ ----------------
        static Font GetFont(string family, float pt, FontStyle style)
        {
            string key = family + "|" + pt + "|" + (int)style;
            Font f;
            if (!fonts.TryGetValue(key, out f)) fonts[key] = f = new Font(family, pt, style);
            return f;
        }

        public static Font Ui(float pt) { return GetFont("Segoe UI", pt, FontStyle.Regular); }
        public static Font UiSemi(float pt) { return GetFont("Segoe UI Semibold", pt, FontStyle.Regular); }
        public static Font Icons(float pt) { return GetFont(IconFamily, pt, FontStyle.Regular); }
        public static Font Mono(float pt) { return GetFont("Consolas", pt, FontStyle.Regular); }

        // ---------------- vẽ ----------------
        public static GraphicsPath Round(RectangleF r, float rad)
        {
            var p = new GraphicsPath();
            float d = Math.Min(rad * 2, Math.Min(r.Width, r.Height));
            if (d < 1) { p.AddRectangle(r); return p; }
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        public static void FillRound(Graphics g, RectangleF r, float rad, Color fill, Color border, float borderWidth)
        {
            var mode = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var path = Round(r, rad))
            {
                if (fill.A > 0) using (var b = new SolidBrush(fill)) g.FillPath(b, path);
                if (border.A > 0 && borderWidth > 0) using (var pen = new Pen(border, borderWidth)) g.DrawPath(pen, path);
            }
            g.SmoothingMode = mode;
        }

        // Biểu tượng từ phông Segoe Fluent Icons / MDL2 (chuỗi thường thì vẽ bằng chữ)
        public static void Glyph(Graphics g, string glyph, RectangleF r, Color c, float pt, bool mirror)
        {
            if (string.IsNullOrEmpty(glyph)) return;
            bool icon = glyph[0] >= '';
            var f = icon ? Icons(pt) : UiSemi(pt * 0.78f);
            var st = g.Save();
            g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
            if (mirror)
            {
                float cx = r.X + r.Width / 2f;
                g.TranslateTransform(cx, 0);
                g.ScaleTransform(-1, 1);
                g.TranslateTransform(-cx, 0);
            }
            using (var b = new SolidBrush(c))
            using (var sf = new StringFormat(StringFormat.GenericTypographic))
            {
                sf.Alignment = StringAlignment.Center;
                sf.LineAlignment = StringAlignment.Center;
                sf.FormatFlags |= StringFormatFlags.NoWrap | StringFormatFlags.NoClip;
                g.DrawString(glyph, f, b, r, sf);
            }
            g.Restore(st);
        }

        // Ô vuông bo góc có màu và biểu tượng trắng (biểu tượng công cụ)
        public static void Badge(Graphics g, RectangleF r, Color fill, string glyph, float pt)
        {
            FillRound(g, r, r.Width * 0.22f, fill, Color.Empty, 0);
            // Một chữ cái (W, X, P như biểu tượng Office) vẽ to hơn chuỗi chữ như "123"
            bool letter = glyph != null && glyph.Length == 1 && glyph[0] < '\uE000';
            Glyph(g, glyph, r, Color.White, letter ? pt * 1.45f : pt, false);
        }

        // ---------------- áp màu cho điều khiển chuẩn ----------------
        public static void Bg(Control c, Role r) { bgRole.Remove(c); bgRole.Add(c, r); }
        public static void Fg(Control c, TextRole r) { fgRole.Remove(c); fgRole.Add(c, r); }

        public static Color Of(Role r)
        {
            switch (r)
            {
                case Role.Surface: return P.Surface;
                case Role.Canvas: return P.Canvas;
                case Role.Input: return P.InputBg;
                case Role.Divider: return P.Divider;
                default: return P.Window;
            }
        }

        public static void Apply(Control root)
        {
            ApplyOne(root);
            foreach (Control c in root.Controls) Apply(c);
        }

        static void ApplyOne(Control c)
        {
            object o;
            Role? role = bgRole.TryGetValue(c, out o) ? (Role?)(Role)o : null;
            var themed = c as IThemed;
            if (themed != null)
            {
                if (role.HasValue) c.BackColor = Of(role.Value);
                themed.ApplyTheme();
                c.Invalidate();
                return;
            }
            var form = c as Form;
            if (form != null) { c.BackColor = P.Window; c.ForeColor = P.Text; Window(form); return; }
            if (c is TextBox)
            {
                if (c.Parent is FieldHost || c.Parent is NumBox) return;   // ô nhập trong khung tự vẽ: khung tự xử lý
                c.BackColor = role.HasValue ? Of(role.Value) : P.InputBg;
                c.ForeColor = P.Text;
                Native(c);
                return;
            }
            if (c is ListView || c is TreeView || c is ListBox)
            {
                c.BackColor = role.HasValue ? Of(role.Value) : P.Surface;
                c.ForeColor = P.Text;
                Native(c);
                c.Invalidate();
                return;
            }
            if (c is Label)
            {
                c.ForeColor = fgRole.TryGetValue(c, out o) ? ((TextRole)o == TextRole.Muted ? P.TextMuted : P.AccentText) : P.Text;
                SetBack(c, role);
                return;
            }
            if (c is PictureBox) { if (role.HasValue) c.BackColor = Of(role.Value); return; }
            c.ForeColor = P.Text;
            if (c is SplitContainer) { c.BackColor = role.HasValue ? Of(role.Value) : P.Window; c.Invalidate(); return; }
            SetBack(c, role);
        }

        static void SetBack(Control c, Role? role)
        {
            if (role.HasValue) { c.BackColor = Of(role.Value); return; }
            try { c.BackColor = Color.Transparent; }
            catch (ArgumentException) { c.BackColor = c.Parent != null ? c.Parent.BackColor : P.Window; }
        }

        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)] static extern int SetWindowTheme(IntPtr h, string app, string id);
        // Hàm không công bố của Windows 10 1903+/11 để điều khiển gốc (cây, danh sách, thanh cuộn) vẽ theo chế độ tối
        [DllImport("uxtheme.dll", EntryPoint = "#135")] static extern int SetPreferredAppMode(int mode);
        [DllImport("uxtheme.dll", EntryPoint = "#133")] static extern bool AllowDarkModeForWindow(IntPtr h, bool allow);
        [DllImport("uxtheme.dll", EntryPoint = "#104")] static extern void RefreshImmersiveColorPolicyState();

        static void AppMode()
        {
            try
            {
                SetPreferredAppMode(P.Dark ? 2 : 3);   // 2 = luôn tối, 3 = luôn sáng
                RefreshImmersiveColorPolicyState();
            }
            catch { }
        }
        [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, IntPtr l);
        [DllImport("dwmapi.dll")] static extern int DwmSetWindowAttribute(IntPtr h, int attr, ref int value, int size);
        [DllImport("user32.dll")] static extern bool RedrawWindow(IntPtr h, IntPtr rect, IntPtr rgn, uint flags);

        // Thanh cuộn, mũi tên, tô sáng của điều khiển Windows theo giao diện sáng/tối
        public static void Native(Control c)
        {
            object o;
            if (!hooked.TryGetValue(c, out o))
            {
                hooked.Add(c, true);
                c.HandleCreated += (s, e) => NativeNow((Control)s);
            }
            if (c.IsHandleCreated) NativeNow(c);
        }

        static void NativeNow(Control c)
        {
            try
            {
                try { AllowDarkModeForWindow(c.Handle, P.Dark); } catch { }
                SetWindowTheme(c.Handle, P.Dark ? "DarkMode_Explorer" : "Explorer", null);
                if (c is ListView)
                {
                    IntPtr header = SendMessage(c.Handle, 0x101F /* LVM_GETHEADER */, IntPtr.Zero, IntPtr.Zero);
                    if (header != IntPtr.Zero)
                    {
                        try { AllowDarkModeForWindow(header, P.Dark); } catch { }
                        SetWindowTheme(header, P.Dark ? "DarkMode_ItemsView" : "ItemsView", null);
                    }
                }
            }
            catch { }
        }

        // Thanh tiêu đề cửa sổ cùng màu với ứng dụng (tối trên Windows 10/11, đổi màu nền trên Windows 11)
        public static void Window(Form f)
        {
            object o;
            if (!hooked.TryGetValue(f, out o))
            {
                hooked.Add(f, true);
                f.HandleCreated += (s, e) => WindowNow((Form)s);
            }
            if (f.IsHandleCreated) WindowNow(f);
        }

        static void WindowNow(Form f)
        {
            try
            {
                int dark = P.Dark ? 1 : 0;
                if (DwmSetWindowAttribute(f.Handle, 20, ref dark, 4) != 0) DwmSetWindowAttribute(f.Handle, 19, ref dark, 4);
                int caption = ColorTranslator.ToWin32(P.Window);
                DwmSetWindowAttribute(f.Handle, 35, ref caption, 4);
                int text = ColorTranslator.ToWin32(P.Text);
                DwmSetWindowAttribute(f.Handle, 36, ref text, 4);
                RedrawWindow(f.Handle, IntPtr.Zero, IntPtr.Zero, 0x0400 | 0x0001 | 0x0004);
            }
            catch { }
        }

        // ---------------- menu và chú thích ----------------
        static Bitmap blank;

        public static ContextMenuStrip Menu()
        {
            return new ContextMenuStrip
            {
                Renderer = new MenuRenderer(), Font = Ui(10f), ShowImageMargin = true,
                ImageScalingSize = new Size(S(16), S(16)), Padding = new Padding(0, S(4), 0, S(4))
            };
        }

        public static ToolStripMenuItem Item(ToolStrip menu, string text, string glyph, EventHandler click)
        {
            var it = new ToolStripMenuItem(text) { Tag = glyph, Padding = new Padding(S(2), S(5), S(2), S(5)) };
            if (glyph != null) it.Image = blank ?? (blank = new Bitmap(16, 16));
            if (click != null) it.Click += click;
            menu.Items.Add(it);
            return it;
        }

        static ToolTip tip;

        public static void Tip(Control c, string text)
        {
            if (tip == null)
            {
                tip = new ToolTip { OwnerDraw = true, InitialDelay = 450, ReshowDelay = 100, AutoPopDelay = 10000 };
                tip.Popup += (s, e) =>
                {
                    var sz = TextRenderer.MeasureText(tip.GetToolTip(e.AssociatedControl), Ui(9f));
                    e.ToolTipSize = new Size(sz.Width + S(18), sz.Height + S(12));
                };
                tip.Draw += (s, e) =>
                {
                    e.Graphics.Clear(P.Menu);
                    var r = e.Bounds;
                    r.Width--; r.Height--;
                    using (var pen = new Pen(P.Border)) e.Graphics.DrawRectangle(pen, r);
                    TextRenderer.DrawText(e.Graphics, e.ToolTipText, Ui(9f), e.Bounds, P.Text,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
                };
            }
            tip.SetToolTip(c, text);
        }

        public const TextFormatFlags TextLeft = TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.NoPrefix |
                                                 TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine | TextFormatFlags.NoPadding;
    }

    // Menu ngữ cảnh theo giao diện sáng/tối
    class MenuRenderer : ToolStripRenderer
    {
        protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e) { e.Graphics.Clear(Theme.P.Menu); }

        protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
        {
            var r = new Rectangle(Point.Empty, e.ToolStrip.Size);
            r.Width--; r.Height--;
            using (var pen = new Pen(Theme.P.Border)) e.Graphics.DrawRectangle(pen, r);
        }

        protected override void OnRenderImageMargin(ToolStripRenderEventArgs e) { }

        protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
        {
            if (!e.Item.Selected || !e.Item.Enabled) return;
            Theme.FillRound(e.Graphics, new RectangleF(Theme.S(4), 1, e.Item.Width - Theme.S(8), e.Item.Height - 2), Theme.S(4), Theme.P.Hover, Color.Empty, 0);
        }

        protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
        {
            e.TextColor = e.Item.Enabled ? Theme.P.Text : Theme.P.TextDisabled;
            base.OnRenderItemText(e);
        }

        protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e)
        {
            int y = e.Item.Height / 2;
            using (var pen = new Pen(Theme.P.Divider)) e.Graphics.DrawLine(pen, Theme.S(8), y, e.Item.Width - Theme.S(8), y);
        }

        // Mục đang chọn: vạch màu nhấn bên trái như Windows 11
        protected override void OnRenderItemCheck(ToolStripItemImageRenderEventArgs e)
        {
            var r = new RectangleF(Theme.S(6), e.Item.Height / 2f - Theme.S(8), Theme.S(3), Theme.S(16));
            Theme.FillRound(e.Graphics, r, Theme.Sf(1.5f), Theme.P.Accent, Color.Empty, 0);
        }

        protected override void OnRenderItemImage(ToolStripItemImageRenderEventArgs e)
        {
            var glyph = e.Item.Tag as string;
            if (glyph != null)
                Theme.Glyph(e.Graphics, glyph, e.ImageRectangle, e.Item.Enabled ? Theme.P.Text : Theme.P.TextDisabled, 10f, false);
        }

        protected override void OnRenderArrow(ToolStripArrowRenderEventArgs e)
        {
            e.ArrowColor = Theme.P.Text;
            base.OnRenderArrow(e);
        }
    }

    // ------------------------------------------------------------------
    // Nút bấm: Standard (viền), Primary (màu nhấn), Subtle (chỉ hiện nền khi rê chuột)
    // ------------------------------------------------------------------
    enum ButtonKind { Standard, Primary, Subtle }

    class FButton : Button, IThemed
    {
        ButtonKind kind;
        string glyph;
        bool showText = true, hover, down;
        public bool MirrorGlyph;
        public float GlyphSize = 11f;
        public bool Compactable;      // được ẩn chữ khi thanh công cụ hẹp

        public FButton(string text, string glyph, ButtonKind kind)
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
            SetStyle(ControlStyles.Opaque, false);   // nút của WinForms mặc định không vẽ nền: tắt để nền trong suốt hoạt động
            BackColor = Color.Transparent;
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            UseVisualStyleBackColor = false;
            Font = Theme.Ui(10f);
            Text = text ?? "";
            this.glyph = glyph;
            this.kind = kind;
            Padding = new Padding(Theme.S(12), 0, Theme.S(12), 0);
            Margin = new Padding(Theme.S(2));
            Size = GetPreferredSize(Size.Empty);
        }

        public ButtonKind Kind { get { return kind; } set { kind = value; Invalidate(); } }
        public string Glyph { get { return glyph; } set { glyph = value; Invalidate(); } }

        public bool ShowText
        {
            get { return showText; }
            set
            {
                if (showText == value) return;
                showText = value;
                if (AutoSize && Parent != null) Parent.PerformLayout();
                else Size = GetPreferredSize(Size.Empty);
                Invalidate();
            }
        }

        bool HasText { get { return showText && !string.IsNullOrEmpty(Text); } }

        // Độ rộng khi hiện đủ chữ (để tính có cần thu gọn thanh công cụ không)
        public int FullWidth
        {
            get
            {
                bool s = showText;
                showText = true;
                int w = GetPreferredSize(Size.Empty).Width;
                showText = s;
                return w;
            }
        }
        bool HasGlyph { get { return !string.IsNullOrEmpty(glyph); } }

        public void ApplyTheme() { Invalidate(); }

        public override Size GetPreferredSize(Size proposedSize)
        {
            int h = Math.Max(Theme.S(32), TextRenderer.MeasureText("Ag", Font).Height + Theme.S(12));
            if (!HasText) return new Size(h, h);
            int w = Padding.Horizontal + TextRenderer.MeasureText(Text, Font, Size.Empty, TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix).Width;
            if (HasGlyph) w += Theme.S(16) + Theme.S(8);
            return new Size(w, h);
        }

        protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hover = down = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { if (e.Button == MouseButtons.Left) { down = true; Invalidate(); } base.OnMouseDown(e); }
        protected override void OnMouseUp(MouseEventArgs e) { down = false; Invalidate(); base.OnMouseUp(e); }
        protected override void OnEnabledChanged(EventArgs e) { hover = down = false; Invalidate(); base.OnEnabledChanged(e); }
        protected override void OnTextChanged(EventArgs e) { base.OnTextChanged(e); Invalidate(); }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            var p = Theme.P;
            Color fill, border = Color.Empty, fg;
            bool en = Enabled;
            if (kind == ButtonKind.Primary)
            {
                fill = !en ? (p.Dark ? Color.FromArgb(0x43, 0x43, 0x43) : Color.FromArgb(0xBF, 0xBF, 0xBF)) : down ? p.AccentPressed : hover ? p.AccentHover : p.Accent;
                fg = en ? p.OnAccent : (p.Dark ? Color.FromArgb(0x8A, 0x8A, 0x8A) : Color.White);
            }
            else if (kind == ButtonKind.Standard)
            {
                fill = !en ? p.ButtonBg : down ? p.ButtonPressed : hover ? p.ButtonHover : p.ButtonBg;
                border = p.ButtonBorder;
                fg = en ? (down ? p.TextMuted : p.Text) : p.TextDisabled;
            }
            else
            {
                fill = en && down ? p.Pressed : en && hover ? p.Hover : Color.Empty;
                fg = en ? p.Text : p.TextDisabled;
            }
            var r = new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f);
            Theme.FillRound(g, r, Theme.S(4), fill, border, 1f);

            int gw = HasGlyph ? Theme.S(16) : 0;
            int gap = HasGlyph && HasText ? Theme.S(8) : 0;
            int tw = HasText ? TextRenderer.MeasureText(g, Text, Font, Size.Empty, TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix).Width : 0;
            int x = (Width - (gw + gap + tw)) / 2;
            if (HasGlyph)
                Theme.Glyph(g, glyph, new RectangleF(x, 0, gw, Height), fg, GlyphSize, MirrorGlyph);
            if (HasText)
                TextRenderer.DrawText(g, Text, Font, new Rectangle(x + gw + gap, 0, tw + 2, Height), fg, Theme.TextLeft);
            if (Focused && ShowFocusCues)
                Theme.FillRound(g, new RectangleF(1.5f, 1.5f, Width - 3f, Height - 3f), Theme.S(4), Color.Empty, p.Text, 1.5f);
        }
    }

    // ------------------------------------------------------------------
    // Ô đánh dấu và nút chọn kiểu Windows 11
    // ------------------------------------------------------------------
    class FCheck : CheckBox, IThemed
    {
        bool hover;

        public FCheck(string text, bool value)
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
            SetStyle(ControlStyles.Opaque, false);
            BackColor = Color.Transparent;
            Font = Theme.Ui(10f);
            Text = text;
            Checked = value;
            AutoSize = true;
        }

        public void ApplyTheme() { Invalidate(); }

        public override Size GetPreferredSize(Size proposedSize)
        {
            var t = TextRenderer.MeasureText(Text, Font, Size.Empty, TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix);
            return new Size(Theme.S(20) + Theme.S(8) + t.Width + Theme.S(4), Math.Max(Theme.S(26), t.Height + Theme.S(6)));
        }

        protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            var p = Theme.P;
            int bs = Theme.S(20);
            var box = new RectangleF(0.5f, (Height - bs) / 2f + 0.5f, bs - 1, bs - 1);
            if (Checked)
            {
                Theme.FillRound(g, box, Theme.S(4), Enabled ? (hover ? p.AccentHover : p.Accent) : p.TextDisabled, Color.Empty, 0);
                Theme.Glyph(g, "", box, p.OnAccent, 9f, false);
            }
            else
                Theme.FillRound(g, box, Theme.S(4), hover && Enabled ? p.ButtonHover : p.InputBg, Enabled ? p.InputBottom : p.TextDisabled, 1f);
            TextRenderer.DrawText(g, Text, Font, new Rectangle(bs + Theme.S(8), 0, Width - bs - Theme.S(8), Height),
                                  Enabled ? p.Text : p.TextDisabled, Theme.TextLeft);
            if (Focused && ShowFocusCues)
                Theme.FillRound(g, RectangleF.Inflate(box, 2, 2), Theme.S(5), Color.Empty, p.Text, 1f);
        }
    }

    class FRadio : RadioButton, IThemed
    {
        bool hover;

        public FRadio(string text, bool value)
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
            SetStyle(ControlStyles.Opaque, false);
            BackColor = Color.Transparent;
            Font = Theme.Ui(10f);
            Text = text;
            Checked = value;
            AutoSize = true;
        }

        public void ApplyTheme() { Invalidate(); }

        public override Size GetPreferredSize(Size proposedSize)
        {
            var t = TextRenderer.MeasureText(Text, Font, Size.Empty, TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix);
            return new Size(Theme.S(20) + Theme.S(8) + t.Width + Theme.S(4), Math.Max(Theme.S(26), t.Height + Theme.S(6)));
        }

        protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnCheckedChanged(EventArgs e) { Invalidate(); base.OnCheckedChanged(e); }
        protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            var p = Theme.P;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            int bs = Theme.S(20);
            var c = new RectangleF(0.5f, (Height - bs) / 2f + 0.5f, bs - 1, bs - 1);
            if (Checked)
            {
                using (var b = new SolidBrush(Enabled ? (hover ? p.AccentHover : p.Accent) : p.TextDisabled)) g.FillEllipse(b, c);
                float d = bs * (hover ? 0.46f : 0.4f);
                using (var b = new SolidBrush(p.OnAccent)) g.FillEllipse(b, c.X + (c.Width - d) / 2, c.Y + (c.Height - d) / 2, d, d);
            }
            else
            {
                using (var b = new SolidBrush(hover && Enabled ? p.ButtonHover : p.InputBg)) g.FillEllipse(b, c);
                using (var pen = new Pen(Enabled ? p.InputBottom : p.TextDisabled)) g.DrawEllipse(pen, c);
            }
            TextRenderer.DrawText(g, Text, Font, new Rectangle(bs + Theme.S(8), 0, Width - bs - Theme.S(8), Height),
                                  Enabled ? p.Text : p.TextDisabled, Theme.TextLeft);
            if (Focused && ShowFocusCues)
                using (var pen = new Pen(p.Text)) g.DrawEllipse(pen, RectangleF.Inflate(c, 2, 2));
        }
    }

    // ------------------------------------------------------------------
    // Khung ô nhập chữ: viền bo góc, vạch màu nhấn bên dưới khi đang nhập
    // ------------------------------------------------------------------
    class FieldHost : Panel, IThemed
    {
        public readonly TextBox Box;
        public string LeadingGlyph;

        public FieldHost(TextBox box)
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.Transparent;
            Box = box;
            box.BorderStyle = BorderStyle.None;
            box.Font = Theme.Ui(10f);
            Controls.Add(box);
            Height = box.PreferredHeight + Theme.S(16);
            Margin = new Padding(Theme.S(3));
            Cursor = Cursors.IBeam;
            box.GotFocus += (s, e) => Invalidate();
            box.LostFocus += (s, e) => Invalidate();
            box.EnabledChanged += (s, e) => ApplyTheme();
            Click += (s, e) => box.Focus();
            ApplyTheme();
        }

        public void ApplyTheme()
        {
            Box.BackColor = Theme.P.InputBg;
            Box.ForeColor = Box.Enabled ? Theme.P.Text : Theme.P.TextDisabled;
            Invalidate();
        }

        // Độ rộng ưa thích nhỏ để bảng bố cục co lại được khi cửa sổ hẹp (ô vẫn giãn theo neo trái-phải)
        public override Size GetPreferredSize(Size proposedSize) { return new Size(Theme.S(120), Box.PreferredHeight + Theme.S(16)); }

        protected override void OnLayout(LayoutEventArgs levent)
        {
            base.OnLayout(levent);
            int left = Theme.S(10) + (LeadingGlyph != null ? Theme.S(24) : 0);
            Box.SetBounds(left, (Height - Box.Height) / 2, Math.Max(Theme.S(20), Width - left - Theme.S(10)), Box.Height);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            DrawFrame(e.Graphics, this, Box.Focused, Box.Enabled);
            if (LeadingGlyph != null)
                Theme.Glyph(e.Graphics, LeadingGlyph, new RectangleF(Theme.S(8), 0, Theme.S(20), Height), Theme.P.TextMuted, 10f, false);
        }

        public static void DrawFrame(Graphics g, Control c, bool focused, bool enabled)
        {
            var p = Theme.P;
            var r = new RectangleF(0.5f, 0.5f, c.Width - 1f, c.Height - 1f);
            Theme.FillRound(g, r, Theme.S(4), p.InputBg, enabled ? p.InputBorder : p.Divider, 1f);
            if (!enabled) return;
            // Vạch dưới (đậm màu nhấn khi đang nhập)
            var st = g.Save();
            float lw = focused ? Theme.S(2) : 1;
            g.SetClip(new RectangleF(0, c.Height - lw - 0.5f, c.Width, lw + 1));
            Theme.FillRound(g, r, Theme.S(4), Color.Empty, focused ? p.Accent : p.InputBottom, focused ? lw * 2 : 1f);
            g.Restore(st);
        }
    }

    // Ô nhập số có nút − / +
    class NumBox : Control, IThemed
    {
        readonly TextBox box = new TextBox { BorderStyle = BorderStyle.None };
        decimal min, max = 1000000, value;
        int hoverBtn = -1;
        public event EventHandler ValueChanged;

        public NumBox()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.Transparent;
            Font = Theme.Ui(10f);
            box.Font = Font;
            box.Text = "0";
            Controls.Add(box);
            Size = new Size(Theme.S(140), box.PreferredHeight + Theme.S(16));
            Margin = new Padding(Theme.S(3));
            box.KeyPress += (s, e) => { if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar)) e.Handled = true; };
            box.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Up) { Step(1); e.Handled = true; }
                else if (e.KeyCode == Keys.Down) { Step(-1); e.Handled = true; }
            };
            box.MouseWheel += (s, e) => Step(e.Delta > 0 ? 1 : -1);
            box.Leave += (s, e) => Commit();
            box.GotFocus += (s, e) => Invalidate();
            box.LostFocus += (s, e) => Invalidate();
            ApplyTheme();
        }

        public decimal Minimum { get { return min; } set { min = value; if (this.value < min) Value = min; } }
        public decimal Maximum { get { return max; } set { max = value; if (this.value > max) Value = max; } }

        public decimal Value
        {
            get { Commit(); return value; }
            set
            {
                decimal v = Math.Max(min, Math.Min(max, value));
                bool changed = v != this.value;
                this.value = v;
                box.Text = v.ToString();
                if (changed) { var h = ValueChanged; if (h != null) h(this, EventArgs.Empty); }
            }
        }

        void Commit()
        {
            decimal v;
            if (!decimal.TryParse(box.Text, out v)) v = value;
            v = Math.Max(min, Math.Min(max, v));
            if (box.Text != v.ToString()) box.Text = v.ToString();
            if (v != value) { value = v; var h = ValueChanged; if (h != null) h(this, EventArgs.Empty); }
        }

        void Step(int d) { Commit(); Value = value + d; box.SelectionStart = box.Text.Length; }

        public void ApplyTheme()
        {
            box.BackColor = Theme.P.InputBg;
            box.ForeColor = Enabled ? Theme.P.Text : Theme.P.TextDisabled;
            Invalidate();
        }

        Rectangle BtnRect(int i)
        {
            int w = Theme.S(28), h = Height - Theme.S(8);
            return new Rectangle(Width - Theme.S(4) - (2 - i) * w, Theme.S(4), w, h);
        }

        protected override void OnLayout(LayoutEventArgs levent)
        {
            base.OnLayout(levent);
            box.SetBounds(Theme.S(10), (Height - box.Height) / 2, Math.Max(Theme.S(20), BtnRect(0).Left - Theme.S(14)), box.Height);
        }

        protected override void OnEnabledChanged(EventArgs e) { base.OnEnabledChanged(e); box.Enabled = Enabled; ApplyTheme(); }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            int h = BtnRect(0).Contains(e.Location) ? 0 : BtnRect(1).Contains(e.Location) ? 1 : -1;
            if (h != hoverBtn) { hoverBtn = h; Invalidate(); }
        }

        protected override void OnMouseLeave(EventArgs e) { base.OnMouseLeave(e); hoverBtn = -1; Invalidate(); }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (!Enabled || e.Button != MouseButtons.Left) return;
            if (BtnRect(0).Contains(e.Location)) Step(-1);
            else if (BtnRect(1).Contains(e.Location)) Step(1);
            else box.Focus();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            FieldHost.DrawFrame(g, this, box.Focused, Enabled);
            for (int i = 0; i < 2; i++)
            {
                var r = BtnRect(i);
                if (i == hoverBtn && Enabled) Theme.FillRound(g, r, Theme.S(3), Theme.P.Hover, Color.Empty, 0);
                Theme.Glyph(g, i == 0 ? "" : "", r, Enabled ? Theme.P.Text : Theme.P.TextDisabled, 8.5f, false);
            }
        }
    }

    // Ô chọn một mục trong danh sách (thay ComboBox để vẽ được theo giao diện tối)
    class FSelect : Control, IThemed
    {
        public readonly List<string> Items = new List<string>();
        int index = -1;
        bool hover, open;
        public event EventHandler SelectedIndexChanged;

        public FSelect()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.Selectable | ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.Transparent;
            Font = Theme.Ui(10f);
            TabStop = true;
            Size = new Size(Theme.S(220), TextRenderer.MeasureText("Ag", Font).Height + Theme.S(16));
            Margin = new Padding(Theme.S(3));
        }

        public override Size GetPreferredSize(Size proposedSize) { return new Size(Theme.S(120), Height); }

        public int SelectedIndex
        {
            get { return index; }
            set
            {
                if (value == index || value < -1 || value >= Items.Count) return;
                index = value;
                Invalidate();
                var h = SelectedIndexChanged;
                if (h != null) h(this, EventArgs.Empty);
            }
        }

        public override string Text
        {
            get { return index >= 0 && index < Items.Count ? Items[index] : ""; }
            set { int i = Items.IndexOf(value); if (i >= 0) SelectedIndex = i; }
        }

        public void ApplyTheme() { Invalidate(); }

        protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
        protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }
        protected override void OnEnabledChanged(EventArgs e) { base.OnEnabledChanged(e); Invalidate(); }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button != MouseButtons.Left || !Enabled) return;
            Focus();
            ShowList();
        }

        protected override bool IsInputKey(Keys keyData)
        {
            Keys k = keyData & Keys.KeyCode;
            return k == Keys.Up || k == Keys.Down || k == Keys.Home || k == Keys.End || base.IsInputKey(keyData);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.KeyCode == Keys.Down && e.Alt || e.KeyCode == Keys.F4 || e.KeyCode == Keys.Space || e.KeyCode == Keys.Enter) { ShowList(); return; }
            if (e.KeyCode == Keys.Up) SelectedIndex = Math.Max(0, index - 1);
            else if (e.KeyCode == Keys.Down) SelectedIndex = Math.Min(Items.Count - 1, index + 1);
            else if (e.KeyCode == Keys.Home) SelectedIndex = 0;
            else if (e.KeyCode == Keys.End) SelectedIndex = Items.Count - 1;
        }

        void ShowList()
        {
            if (Items.Count == 0) return;
            var menu = Theme.Menu();
            menu.ShowImageMargin = true;
            for (int i = 0; i < Items.Count; i++)
            {
                int k = i;
                var it = Theme.Item(menu, Items[i], null, (s, e) => SelectedIndex = k);
                it.Checked = i == index;
            }
            menu.MinimumSize = new Size(Width, 0);
            menu.Closed += (s, e) => { open = false; Invalidate(); BeginInvoke(new Action(menu.Dispose)); };
            open = true;
            Invalidate();
            menu.Show(this, new Point(0, Height + Theme.S(2)));
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            var p = Theme.P;
            bool focus = Focused || open;
            var r = new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f);
            Theme.FillRound(g, r, Theme.S(4), Enabled ? (hover ? p.ButtonHover : p.ButtonBg) : p.InputBg, Enabled ? p.ButtonBorder : p.Divider, 1f);
            if (focus && Enabled)
            {
                var st = g.Save();
                g.SetClip(new RectangleF(0, Height - Theme.S(2) - 0.5f, Width, Theme.S(2) + 1));
                Theme.FillRound(g, r, Theme.S(4), Color.Empty, p.Accent, Theme.S(4));
                g.Restore(st);
            }
            TextRenderer.DrawText(g, Text, Font, new Rectangle(Theme.S(11), 0, Width - Theme.S(44), Height), Enabled ? p.Text : p.TextDisabled, Theme.TextLeft);
            Theme.Glyph(g, "", new RectangleF(Width - Theme.S(32), 0, Theme.S(22), Height), Enabled ? p.TextMuted : p.TextDisabled, 8f, false);
        }
    }

    // Nút chọn màu (ô màu)
    class ColorSwatch : Control, IThemed
    {
        Color value;
        bool hover;

        public ColorSwatch(Color value)
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.Transparent;
            this.value = value;
            Size = new Size(Theme.S(96), Theme.S(34));
            Margin = new Padding(Theme.S(3));
            Cursor = Cursors.Hand;
        }

        public Color Value { get { return value; } set { this.value = value; Invalidate(); } }
        public void ApplyTheme() { Invalidate(); }

        protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }

        protected override void OnClick(EventArgs e)
        {
            base.OnClick(e);
            using (var d = new ColorDialog { Color = value, FullOpen = true })
                if (d.ShowDialog(FindForm()) == DialogResult.OK) Value = d.Color;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var p = Theme.P;
            Theme.FillRound(e.Graphics, new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f), Theme.S(4), hover ? p.ButtonHover : p.ButtonBg, p.ButtonBorder, 1f);
            Theme.FillRound(e.Graphics, new RectangleF(Theme.S(5), Theme.S(5), Width - Theme.S(30), Height - Theme.S(10)), Theme.S(3), value, Color.FromArgb(40, 0, 0, 0), 1f);
            Theme.Glyph(e.Graphics, "", new RectangleF(Width - Theme.S(24), 0, Theme.S(20), Height), p.TextMuted, 8f, false);
        }
    }

    // ------------------------------------------------------------------
    // Thẻ (card): nền bo góc, có tiêu đề và biểu tượng
    // ------------------------------------------------------------------
    class Card : Panel, IThemed
    {
        string title;
        public string Glyph;

        public Card(string title, string glyph)
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.Transparent;
            this.title = title;
            Glyph = glyph;
            Padding = new Padding(Theme.S(16), title != null ? Theme.S(48) : Theme.S(14), Theme.S(16), Theme.S(14));
            Margin = new Padding(Theme.S(6));
        }

        public string Title { get { return title; } set { title = value; Invalidate(); } }
        public void ApplyTheme() { Invalidate(); }

        // Chỉ báo chiều cao cần thiết; độ rộng lấy theo ô chứa (để bảng bố cục co lại được khi cửa sổ hẹp)
        public override Size GetPreferredSize(Size proposedSize)
        {
            Size s = base.GetPreferredSize(proposedSize);
            return new Size(Math.Min(s.Width, Theme.S(160)), s.Height);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            var p = Theme.P;
            Theme.FillRound(g, new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f), Theme.S(8), p.Surface, p.Border, 1f);
            if (title == null) return;
            int x = Theme.S(16);
            if (Glyph != null)
            {
                Theme.Glyph(g, Glyph, new RectangleF(x, Theme.S(14), Theme.S(20), Theme.S(22)), p.AccentText, 11f, false);
                x += Theme.S(30);
            }
            TextRenderer.DrawText(g, title, Theme.UiSemi(11f), new Rectangle(x, Theme.S(13), Width - x - Theme.S(16), Theme.S(24)), p.Text, Theme.TextLeft);
        }
    }

    // Hàng nút tự xuống dòng khi hẹp; không ép bảng bố cục giãn rộng theo tổng độ rộng các nút
    class WrapFlow : FlowLayoutPanel
    {
        public WrapFlow() { WrapContents = true; }

        public override Size GetPreferredSize(Size proposedSize)
        {
            int w = Width > 0 ? Width : proposedSize.Width;
            Size s = base.GetPreferredSize(new Size(Math.Max(1, w), 0));
            return new Size(Math.Min(s.Width, Theme.S(160)), s.Height);
        }
    }

    // Thanh tiến trình mảnh chạy liên tục (không xác định thời gian)
    class ThinProgress : Control
    {
        readonly System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer { Interval = 16 };
        float t = -0.4f;

        public ThinProgress()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.Transparent;
            Size = new Size(Theme.S(160), Theme.S(4));
            timer.Tick += (s, e) => { t += 0.012f; if (t > 1.2f) t = -0.4f; Invalidate(); };
        }

        protected override void OnVisibleChanged(EventArgs e) { base.OnVisibleChanged(e); timer.Enabled = Visible; }

        protected override void Dispose(bool disposing) { if (disposing) timer.Dispose(); base.Dispose(disposing); }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            float h = Height;
            Theme.FillRound(g, new RectangleF(0, 0, Width, h), h / 2, Theme.P.Divider, Color.Empty, 0);
            float x = t * Width, w = Width * 0.4f;
            float x0 = Math.Max(0, x), x1 = Math.Min(Width, x + w);
            if (x1 > x0) Theme.FillRound(g, new RectangleF(x0, 0, x1 - x0, h), h / 2, Theme.P.Accent, Color.Empty, 0);
        }
    }

    // Chọn giữa các trang chính (Soạn PDF / Hộp công cụ)
    class NavSegment : Control, IThemed
    {
        readonly List<string[]> items = new List<string[]>();
        int index, hover = -1;
        public event EventHandler SelectedIndexChanged;

        public NavSegment()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.Transparent;
            Font = Theme.Ui(10f);
            Height = Theme.S(40);
        }

        public void Add(string glyph, string text)
        {
            items.Add(new[] { glyph, text });
            Width = Theme.S(8) + Enumerable.Range(0, items.Count).Sum(i => ItemWidth(i));
        }

        int ItemWidth(int i) { return Theme.S(16 + 8 + 32) + TextRenderer.MeasureText(items[i][1], Font).Width; }

        Rectangle ItemRect(int i)
        {
            int x = Theme.S(4);
            for (int k = 0; k < i; k++) x += ItemWidth(k);
            return new Rectangle(x, Theme.S(4), ItemWidth(i), Height - Theme.S(8));
        }

        public int SelectedIndex
        {
            get { return index; }
            set
            {
                if (value == index || value < 0 || value >= items.Count) return;
                index = value;
                Invalidate();
                var h = SelectedIndexChanged;
                if (h != null) h(this, EventArgs.Empty);
            }
        }

        public void ApplyTheme() { Invalidate(); }

        int Hit(Point pt)
        {
            for (int i = 0; i < items.Count; i++) if (ItemRect(i).Contains(pt)) return i;
            return -1;
        }

        protected override void OnMouseMove(MouseEventArgs e) { base.OnMouseMove(e); int h = Hit(e.Location); if (h != hover) { hover = h; Invalidate(); } }
        protected override void OnMouseLeave(EventArgs e) { base.OnMouseLeave(e); hover = -1; Invalidate(); }
        protected override void OnMouseDown(MouseEventArgs e) { base.OnMouseDown(e); int h = Hit(e.Location); if (h >= 0) SelectedIndex = h; }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            var p = Theme.P;
            Theme.FillRound(g, new RectangleF(0.5f, 0.5f, Width - 1f, Height - 1f), Theme.S(8),
                            p.Dark ? Color.FromArgb(255, 0x2A, 0x2A, 0x2A) : Color.FromArgb(255, 0xE7, 0xE7, 0xE7), Color.Empty, 0);
            for (int i = 0; i < items.Count; i++)
            {
                var r = ItemRect(i);
                bool sel = i == index;
                if (sel)
                    Theme.FillRound(g, new RectangleF(r.X + 0.5f, r.Y + 0.5f, r.Width - 1f, r.Height - 1f), Theme.S(6),
                                    p.Dark ? Color.FromArgb(255, 0x3F, 0x3F, 0x3F) : Color.White, p.Dark ? Color.Empty : Color.FromArgb(255, 0xDE, 0xDE, 0xDE), 1f);
                else if (i == hover)
                    Theme.FillRound(g, r, Theme.S(6), p.Hover, Color.Empty, 0);
                Color fg = sel || i == hover ? p.Text : p.TextMuted;
                int tw = TextRenderer.MeasureText(items[i][1], Font).Width;
                int x = r.X + (r.Width - (Theme.S(24) + tw)) / 2;
                Theme.Glyph(g, items[i][0], new RectangleF(x, r.Y, Theme.S(16), r.Height), sel ? p.AccentText : fg, 11f, false);
                TextRenderer.DrawText(g, items[i][1], Font, new Rectangle(x + Theme.S(24), r.Y, tw + 4, r.Height), fg, Theme.TextLeft);
            }
        }
    }

    // Vùng gợi ý kéo thả khi danh sách file còn trống
    class DropHint : Control, IThemed
    {
        public string Title = "", Sub = "";
        public string Glyph = "";
        bool hover;

        public DropHint()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.Transparent;
            Cursor = Cursors.Hand;
        }

        public void ApplyTheme() { Invalidate(); }
        protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            var p = Theme.P;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var r = new RectangleF(Theme.S(2), Theme.S(2), Width - Theme.S(4), Height - Theme.S(4));
            using (var path = Theme.Round(r, Theme.S(8)))
            {
                if (hover) using (var b = new SolidBrush(p.Hover)) g.FillPath(b, path);
                using (var pen = new Pen(p.Dark ? Color.FromArgb(255, 0x55, 0x55, 0x55) : Color.FromArgb(255, 0xC4, 0xC4, 0xC4), 1.5f) { DashStyle = DashStyle.Dash })
                    g.DrawPath(pen, path);
            }
            int cy = Height / 2;
            var circle = new RectangleF(Width / 2f - Theme.S(28), cy - Theme.S(66), Theme.S(56), Theme.S(56));
            using (var b = new SolidBrush(Color.FromArgb(p.Dark ? 40 : 24, p.Accent))) g.FillEllipse(b, circle);
            Theme.Glyph(g, Glyph, circle, p.AccentText, 20f, false);
            TextRenderer.DrawText(g, Title, Theme.UiSemi(11f), new Rectangle(0, cy + Theme.S(2), Width, Theme.S(26)), p.Text,
                                  TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
            TextRenderer.DrawText(g, Sub, Theme.Ui(9.5f), new Rectangle(0, cy + Theme.S(30), Width, Theme.S(22)), p.TextMuted,
                                  TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
        }
    }

    // Ô biểu tượng màu của công cụ (đầu trang Hộp công cụ)
    class IconBadge : Control, IThemed
    {
        public Color Fill = Color.Gray;
        public string Glyph = "";

        public IconBadge()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.Transparent;
        }

        public void ApplyTheme() { Invalidate(); }

        protected override void OnPaint(PaintEventArgs e)
        {
            Theme.Badge(e.Graphics, new RectangleF(0, 0, Width - 1, Height - 1), Fill, Glyph, Height * 0.3f / Theme.Scale);
        }
    }

    // Vạch ngăn mảnh (dùng trong thanh công cụ)
    class Separator : Control, IThemed
    {
        public Separator()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.Transparent;
            Size = new Size(Theme.S(9), Theme.S(32));
            Margin = new Padding(Theme.S(2));
        }

        public void ApplyTheme() { Invalidate(); }

        protected override void OnPaint(PaintEventArgs e)
        {
            using (var pen = new Pen(Theme.P.Border)) e.Graphics.DrawLine(pen, Width / 2, Theme.S(6), Width / 2, Height - Theme.S(6));
        }
    }


    // ------------------------------------------------------------------
    // Danh sách file tự vẽ (thay ListView): tiêu đề cột cố định, chọn nhiều dòng,
    // bàn phím, kéo thả, rê chuột... vẽ đúng theo giao diện sáng/tối
    // ------------------------------------------------------------------
    class FileList : Panel, IThemed
    {
        public class Column
        {
            public string Text;
            public int Width;
            public bool Right;
        }

        public class Row
        {
            public string[] Cells;
            public string ImageKey;
            public object Tag;
            public bool Selected;
        }

        public readonly List<Column> Columns = new List<Column>();
        public readonly List<Row> Rows = new List<Row>();
        public ImageList Images;
        public int FillColumn;
        public Func<int, int> SortState;          // -1 không sắp xếp, 0 tăng dần, 1 giảm dần
        public bool ClickableHeader = true;
        public event EventHandler SelectionChanged, ItemActivate;
        public event Action<int> HeaderClick;
        public event Action ItemDragStart;

        int anchor = -1, focusRow = -1, hoverRow = -1, hoverHeader = -1;
        int pendingSingle = -1, resizeCol = -1, resizeStartX, resizeStartW;
        bool dragArmed;
        Point downPt;
        int updating;

        static int S(int v) { return Theme.S(v); }
        int HeaderH { get { return S(32); } }
        int RowH { get { return S(30); } }

        public FileList()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
            AutoScroll = true;
            TabStop = true;
            Font = Theme.Ui(10f);
        }

        public void ApplyTheme()
        {
            BackColor = Theme.P.Surface;
            Theme.Native(this);
            Invalidate();
        }

        // ---------------- dữ liệu ----------------
        public void BeginUpdate() { updating++; }
        public void EndUpdate() { if (--updating <= 0) { updating = 0; Reflow(); } }

        public void Clear()
        {
            Rows.Clear();
            anchor = focusRow = hoverRow = pendingSingle = -1;
            if (updating == 0) Reflow();
        }

        public Row Add(string[] cells, string imageKey, object tag, bool selected)
        {
            var r = new Row { Cells = cells, ImageKey = imageKey, Tag = tag, Selected = selected };
            Rows.Add(r);
            if (updating == 0) Reflow();
            return r;
        }

        public List<int> SelectedIndices { get { return Enumerable.Range(0, Rows.Count).Where(i => Rows[i].Selected).ToList(); } }
        public List<Row> SelectedRows { get { return Rows.Where(r => r.Selected).ToList(); } }

        public int TopIndex
        {
            get { return Math.Max(0, -AutoScrollPosition.Y / RowH); }
            set { AutoScrollPosition = new Point(0, Math.Max(0, value) * RowH); }
        }

        public void SelectOnly(int i)
        {
            for (int k = 0; k < Rows.Count; k++) Rows[k].Selected = k == i;
            anchor = focusRow = i;
            EnsureVisible(i);
            Changed();
        }

        public void SelectAll()
        {
            foreach (var r in Rows) r.Selected = true;
            Changed();
        }

        void Changed()
        {
            Invalidate();
            var h = SelectionChanged;
            if (h != null) h(this, EventArgs.Empty);
        }

        void Reflow()
        {
            FitColumns();
            AutoScrollMinSize = new Size(0, HeaderH + Rows.Count * RowH + S(4));
            if (focusRow >= Rows.Count) focusRow = -1;
            Invalidate();
        }

        void FitColumns()
        {
            if (FillColumn < 0 || FillColumn >= Columns.Count) return;
            int others = 0;
            for (int i = 0; i < Columns.Count; i++) if (i != FillColumn) others += Columns[i].Width;
            Columns[FillColumn].Width = Math.Max(S(120), ClientSize.Width - others - S(8));
        }

        protected override void OnClientSizeChanged(EventArgs e)
        {
            base.OnClientSizeChanged(e);
            FitColumns();
            Invalidate();
        }

        public void EnsureVisible(int i)
        {
            if (i < 0 || i >= Rows.Count) return;
            int y = i * RowH, top = -AutoScrollPosition.Y, visible = ClientSize.Height - HeaderH;
            if (y < top) AutoScrollPosition = new Point(0, y);
            else if (y + RowH > top + visible) AutoScrollPosition = new Point(0, y + RowH - visible);
            Invalidate();
        }

        // ---------------- vị trí ----------------
        int RowAt(Point p)
        {
            if (p.Y < HeaderH) return -1;
            int i = (p.Y - HeaderH - AutoScrollPosition.Y) / RowH;
            return i >= 0 && i < Rows.Count ? i : -1;
        }

        int ColumnLeft(int c)
        {
            int x = S(4);
            for (int k = 0; k < c; k++) x += Columns[k].Width;
            return x;
        }

        int HeaderAt(Point p)
        {
            if (p.Y >= HeaderH) return -1;
            for (int c = 0; c < Columns.Count; c++)
            {
                int x = ColumnLeft(c);
                if (p.X >= x && p.X < x + Columns[c].Width) return c;
            }
            return -1;
        }

        // Cạnh phải của cột (để kéo đổi độ rộng), trả về chỉ số cột hoặc -1
        int DividerAt(Point p)
        {
            if (p.Y >= HeaderH) return -1;
            for (int c = 0; c < Columns.Count; c++)
            {
                if (c == FillColumn) continue;
                int x = ColumnLeft(c) + Columns[c].Width;
                if (Math.Abs(p.X - x) <= S(4)) return c;
            }
            return -1;
        }

        // ---------------- vẽ ----------------
        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            var p = Theme.P;
            g.Clear(p.Surface);
            int top = -AutoScrollPosition.Y;
            int first = Math.Max(0, top / RowH), last = Math.Min(Rows.Count - 1, (top + ClientSize.Height) / RowH + 1);
            var font = Font;
            for (int i = first; i <= last; i++)
            {
                var row = Rows[i];
                var rr = new Rectangle(S(4), HeaderH + i * RowH - top, ClientSize.Width - S(8), RowH);
                if (rr.Bottom < HeaderH) continue;
                var bg = new RectangleF(rr.X, rr.Y + 1, rr.Width, rr.Height - 2);
                if (row.Selected) Theme.FillRound(g, bg, S(4), Focused ? p.Selection : (p.Dark ? Color.FromArgb(255, 0x3A, 0x3A, 0x3A) : Color.FromArgb(255, 0xE8, 0xE8, 0xE8)), Color.Empty, 0);
                else if (i == hoverRow) Theme.FillRound(g, bg, S(4), p.Hover, Color.Empty, 0);
                if (i == focusRow && Focused && ShowFocusCues) Theme.FillRound(g, bg, S(4), Color.Empty, p.Text, 1);
                for (int c = 0; c < Columns.Count && c < row.Cells.Length; c++)
                {
                    int x = ColumnLeft(c), w = Columns[c].Width;
                    var cell = new Rectangle(x + S(8), rr.Y, w - S(14), rr.Height);
                    if (c == 0 && Images != null && row.ImageKey != null && Images.Images.ContainsKey(row.ImageKey))
                    {
                        g.DrawImage(Images.Images[row.ImageKey], cell.X, rr.Y + (rr.Height - S(16)) / 2, S(16), S(16));
                        cell.X += S(24);
                        cell.Width -= S(24);
                    }
                    TextRenderer.DrawText(g, row.Cells[c], font, cell, c == 0 ? p.Text : p.TextMuted,
                        Theme.TextLeft | (Columns[c].Right ? TextFormatFlags.Right : TextFormatFlags.Left));
                }
            }
            DrawHeader(g);
        }

        void DrawHeader(Graphics g)
        {
            var p = Theme.P;
            using (var b = new SolidBrush(p.Surface)) g.FillRectangle(b, 0, 0, ClientSize.Width, HeaderH);
            using (var pen = new Pen(p.Divider)) g.DrawLine(pen, 0, HeaderH - 1, ClientSize.Width, HeaderH - 1);
            var f = Theme.Ui(9.5f);
            for (int c = 0; c < Columns.Count; c++)
            {
                int x = ColumnLeft(c), w = Columns[c].Width;
                if (ClickableHeader && c == hoverHeader)
                    Theme.FillRound(g, new RectangleF(x + 1, S(3), w - 2, HeaderH - S(6)), S(4), p.Hover, Color.Empty, 0);
                if (c < Columns.Count - 1)
                    using (var pen = new Pen(p.Divider)) g.DrawLine(pen, x + w - 1, S(8), x + w - 1, HeaderH - S(8));
                var tr = new Rectangle(x + S(8), 0, w - S(14), HeaderH);
                int sort = SortState != null ? SortState(c) : -1;
                if (sort >= 0)
                {
                    var ar = Columns[c].Right ? new RectangleF(tr.X, 0, S(14), HeaderH) : new RectangleF(tr.Right - S(14), 0, S(14), HeaderH);
                    Theme.Glyph(g, sort == 0 ? "" : "", ar, p.AccentText, 7.5f, false);
                    if (Columns[c].Right) tr.X += S(16);
                    tr.Width -= S(16);
                }
                TextRenderer.DrawText(g, Columns[c].Text, f, tr, sort >= 0 ? p.Text : p.TextMuted,
                    Theme.TextLeft | (Columns[c].Right ? TextFormatFlags.Right : TextFormatFlags.Left));
            }
        }

        // ---------------- chuột ----------------
        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            Focus();
            if (e.Y < HeaderH)
            {
                if (e.Button != MouseButtons.Left) return;
                int d = DividerAt(e.Location);
                if (d >= 0) { resizeCol = d; resizeStartX = e.X; resizeStartW = Columns[d].Width; return; }
                int c = HeaderAt(e.Location);
                if (c >= 0 && ClickableHeader && HeaderClick != null) HeaderClick(c);
                return;
            }
            int i = RowAt(e.Location);
            bool ctrl = (ModifierKeys & Keys.Control) != 0, shift = (ModifierKeys & Keys.Shift) != 0;
            pendingSingle = -1;
            dragArmed = false;
            if (i < 0)
            {
                if (!ctrl && Rows.Any(r => r.Selected)) { foreach (var r in Rows) r.Selected = false; Changed(); }
                return;
            }
            if (e.Button == MouseButtons.Right)
            {
                if (!Rows[i].Selected) SelectOnly(i);
                return;
            }
            if (e.Clicks == 2) { var h = ItemActivate; if (h != null) h(this, EventArgs.Empty); return; }
            if (shift && anchor >= 0 && anchor < Rows.Count)
            {
                if (!ctrl) foreach (var r in Rows) r.Selected = false;
                for (int k = Math.Min(anchor, i); k <= Math.Max(anchor, i); k++) Rows[k].Selected = true;
                focusRow = i;
            }
            else if (ctrl)
            {
                Rows[i].Selected = !Rows[i].Selected;
                anchor = focusRow = i;
            }
            else if (Rows[i].Selected && Rows.Count(r => r.Selected) > 1)
            {
                pendingSingle = i;
                focusRow = i;
            }
            else
            {
                foreach (var r in Rows) r.Selected = false;
                Rows[i].Selected = true;
                anchor = focusRow = i;
            }
            dragArmed = Rows[i].Selected;
            downPt = e.Location;
            Changed();
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (resizeCol >= 0)
            {
                Columns[resizeCol].Width = Math.Max(S(50), resizeStartW + e.X - resizeStartX);
                FitColumns();
                Invalidate();
                return;
            }
            if (e.Button == MouseButtons.None)
            {
                Cursor = DividerAt(e.Location) >= 0 ? Cursors.VSplit : Cursors.Default;
                int h = RowAt(e.Location), hh = ClickableHeader ? HeaderAt(e.Location) : -1;
                if (h != hoverRow || hh != hoverHeader) { hoverRow = h; hoverHeader = hh; Invalidate(); }
                return;
            }
            if (!dragArmed || e.Button != MouseButtons.Left) return;
            var ds = SystemInformation.DragSize;
            if (Math.Abs(e.X - downPt.X) < ds.Width && Math.Abs(e.Y - downPt.Y) < ds.Height) return;
            dragArmed = false;
            pendingSingle = -1;
            if (ItemDragStart != null) ItemDragStart();
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            resizeCol = -1;
            dragArmed = false;
            if (pendingSingle >= 0 && pendingSingle < Rows.Count)
            {
                int i = pendingSingle;
                pendingSingle = -1;
                foreach (var r in Rows) r.Selected = false;
                Rows[i].Selected = true;
                anchor = focusRow = i;
                Changed();
            }
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            if (hoverRow < 0 && hoverHeader < 0) return;
            hoverRow = hoverHeader = -1;
            Invalidate();
        }

        protected override void OnGotFocus(EventArgs e) { base.OnGotFocus(e); Invalidate(); }
        protected override void OnLostFocus(EventArgs e) { base.OnLostFocus(e); Invalidate(); }

        protected override void OnScroll(ScrollEventArgs se) { base.OnScroll(se); Invalidate(); }
        protected override void OnMouseWheel(MouseEventArgs e) { base.OnMouseWheel(e); Invalidate(); }

        // ---------------- bàn phím ----------------
        protected override bool IsInputKey(Keys keyData)
        {
            Keys k = keyData & Keys.KeyCode;
            return k == Keys.Up || k == Keys.Down || k == Keys.Home || k == Keys.End || k == Keys.PageUp || k == Keys.PageDown || base.IsInputKey(keyData);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (Rows.Count > 0)
            {
                int page = Math.Max(1, (ClientSize.Height - HeaderH) / RowH - 1);
                int target = -1;
                switch (e.KeyCode)
                {
                    case Keys.Up: target = Math.Max(0, focusRow - 1); break;
                    case Keys.Down: target = Math.Min(Rows.Count - 1, focusRow + 1); break;
                    case Keys.Home: target = 0; break;
                    case Keys.End: target = Rows.Count - 1; break;
                    case Keys.PageUp: target = Math.Max(0, focusRow - page); break;
                    case Keys.PageDown: target = Math.Min(Rows.Count - 1, focusRow + page); break;
                }
                if (target >= 0 && !e.Alt)
                {
                    if (e.Shift && anchor >= 0)
                    {
                        foreach (var r in Rows) r.Selected = false;
                        for (int k = Math.Min(anchor, target); k <= Math.Max(anchor, target); k++) Rows[k].Selected = true;
                        focusRow = target;
                        EnsureVisible(target);
                        Changed();
                    }
                    else SelectOnly(target);
                    e.Handled = true;
                    return;
                }
                if (e.Control && e.KeyCode == Keys.A) { SelectAll(); e.Handled = true; return; }
                if (e.KeyCode == Keys.Enter && Rows.Any(r => r.Selected))
                {
                    e.SuppressKeyPress = true;
                    var h = ItemActivate;
                    if (h != null) h(this, EventArgs.Empty);
                    return;
                }
            }
            base.OnKeyDown(e);
        }
    }

    // TreeView không có thanh cuộn ngang (tên dài thì cắt "...")
    class NavTree : TreeView
    {
        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.Style |= 0x8000;   // TVS_NOHSCROLL
                return cp;
            }
        }
    }

    // Cây thư mục tự vẽ toàn bộ (giữ phần mở rộng/cuộn/bàn phím của TreeView gốc)
    static class TreeStyle
    {
        [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, IntPtr l);

        public static void Setup(TreeView t, object separator)
        {
            t.DrawMode = TreeViewDrawMode.OwnerDrawAll;
            t.HotTracking = true;
            t.ShowLines = false;
            t.FullRowSelect = true;
            // Vẽ đệm kép để không nháy (TVS_EX_DOUBLEBUFFER)
            t.HandleCreated += (s, e) => SendMessage(t.Handle, 0x112C, (IntPtr)0x0004, (IntPtr)0x0004);
            t.DrawNode += (s, e) =>
            {
                var g = e.Graphics;
                var p = Theme.P;
                var row = new Rectangle(0, e.Bounds.Y, t.ClientSize.Width, e.Bounds.Height);
                using (var b = new SolidBrush(t.BackColor)) g.FillRectangle(b, row);
                if (e.Node.Tag == separator)
                {
                    int y = row.Y + row.Height / 2;
                    using (var pen = new Pen(p.Divider)) g.DrawLine(pen, Theme.S(10), y, row.Right - Theme.S(10), y);
                    return;
                }
                bool sel = (e.State & TreeNodeStates.Selected) != 0;
                bool hot = (e.State & TreeNodeStates.Hot) != 0;
                var bg = new RectangleF(row.X + Theme.S(4), row.Y + 1, row.Width - Theme.S(8), row.Height - 2);
                if (sel)
                {
                    Theme.FillRound(g, bg, Theme.S(4), p.Dark ? Color.FromArgb(255, 0x3A, 0x3A, 0x3A) : Color.FromArgb(255, 0xE6, 0xE6, 0xE6), Color.Empty, 0);
                    Theme.FillRound(g, new RectangleF(bg.X, bg.Y + bg.Height / 2 - Theme.S(8), Theme.S(3), Theme.S(16)), Theme.Sf(1.5f), p.Accent, Color.Empty, 0);
                }
                else if (hot) Theme.FillRound(g, bg, Theme.S(4), p.Hover, Color.Empty, 0);

                var text = e.Node.Bounds;
                int iw = t.ImageList != null ? t.ImageList.ImageSize.Width : 0;
                int imgX = text.X - iw - 3;
                if (e.Node.Nodes.Count > 0)
                    Theme.Glyph(g, e.Node.IsExpanded ? "" : "", new RectangleF(imgX - Theme.S(16), row.Y, Theme.S(16), row.Height), p.TextMuted, 7f, false);
                if (iw > 0 && e.Node.ImageKey != null && t.ImageList.Images.ContainsKey(e.Node.ImageKey))
                    g.DrawImage(t.ImageList.Images[e.Node.ImageKey], imgX, row.Y + (row.Height - iw) / 2, iw, iw);
                TextRenderer.DrawText(g, e.Node.Text, t.Font, new Rectangle(text.X + Theme.S(2), row.Y, Math.Max(10, row.Right - text.X - Theme.S(10)), row.Height),
                                      p.Text, Theme.TextLeft);
            };
        }
    }
}
