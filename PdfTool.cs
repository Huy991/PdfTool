// PdfTool - bộ công cụ PDF (gộp, tách, nén, xoay, chuyển đổi, OCR, đánh số trang, đóng dấu, mật khẩu)
// Dùng các công cụ dòng lệnh mã nguồn mở: qpdf, Ghostscript, Tesseract (tìm trong thư mục tools\ cạnh exe,
// hoặc trong thư mục cài đặt PDF24, hoặc trong PATH).
// Biên dịch: chạy build.bat (dùng csc.exe có sẵn của .NET Framework 4.x)

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace PdfTool
{
    // ------------------------------------------------------------------
    // Tìm và chạy các công cụ dòng lệnh bên ngoài
    // ------------------------------------------------------------------
    public static class Ext
    {
        static string[] Roots()
        {
            var list = new List<string>();
            list.Add(Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "tools"));
            foreach (var pf in new[] { Environment.GetEnvironmentVariable("ProgramW6432"), Environment.GetEnvironmentVariable("ProgramFiles"), Environment.GetEnvironmentVariable("ProgramFiles(x86)") })
                if (!string.IsNullOrEmpty(pf)) list.Add(Path.Combine(pf, "PDF24"));
            return list.ToArray();
        }

        static string Find(string[] relative, string exeName)
        {
            foreach (string root in Roots())
                foreach (string rel in relative)
                {
                    string p = Path.Combine(root, rel);
                    if (File.Exists(p)) return p;
                }
            // Ghostscript cài riêng: C:\Program Files\gs\gs10.xx\bin
            if (exeName.StartsWith("gswin"))
            {
                string gsRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "gs");
                if (Directory.Exists(gsRoot))
                    foreach (string d in Directory.GetDirectories(gsRoot).OrderByDescending(x => x))
                    {
                        string p = Path.Combine(d, "bin", exeName);
                        if (File.Exists(p)) return p;
                    }
            }
            foreach (string dir in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(';'))
            {
                try
                {
                    string p = Path.Combine(dir.Trim(), exeName);
                    if (dir.Trim().Length > 0 && File.Exists(p)) return p;
                }
                catch { }
            }
            return null;
        }

        static string qpdf, gs, tess;
        public static string Qpdf { get { return qpdf ?? (qpdf = Find(new[] { @"qpdf\bin\qpdf.exe", "qpdf.exe" }, "qpdf.exe")); } }
        public static string Gs { get { return gs ?? (gs = Find(new[] { @"gs\bin\gswin64c.exe", @"gs\bin\gswinc.exe", "gswin64c.exe" }, "gswin64c.exe")); } }
        public static string Tesseract { get { return tess ?? (tess = Find(new[] { @"tesseract\tesseract.exe", "tesseract.exe" }, "tesseract.exe")); } }

        public static string Require(string path, string name)
        {
            if (path == null)
                throw new InvalidOperationException("Không tìm thấy " + name + ". Hãy cài PDF24 hoặc chép " + name + " vào thư mục \"tools\" cạnh PdfTool.exe.");
            return path;
        }

        // Ghép tham số theo quy tắc của Windows (CommandLineToArgvW)
        public static string Q(string s)
        {
            if (s.Length > 0 && s.IndexOfAny(new[] { ' ', '\t', '"', '&', '|', '<', '>', '^' }) < 0) return s;
            var sb = new StringBuilder("\"");
            int bs = 0;
            foreach (char c in s)
            {
                if (c == '\\') { bs++; continue; }
                if (c == '"') { sb.Append('\\', bs * 2 + 1); sb.Append('"'); }
                else { sb.Append('\\', bs); sb.Append(c); }
                bs = 0;
            }
            sb.Append('\\', bs * 2);
            sb.Append('"');
            return sb.ToString();
        }

        // Tiến trình đang chạy theo từng luồng, để "Dừng" chỉ hủy đúng công việc của luồng đó
        static readonly Dictionary<Thread, Process> running = new Dictionary<Thread, Process>();
        public static void Kill(Thread t)
        {
            Process p;
            lock (running) running.TryGetValue(t, out p);
            try { if (p != null && !p.HasExited) p.Kill(); } catch { }
        }

        public static int Run(string exe, string args, out string output)
        {
            var psi = new ProcessStartInfo(exe, args)
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8,
                WorkingDirectory = Path.GetDirectoryName(exe)
            };
            var sb = new StringBuilder();
            using (var p = new Process { StartInfo = psi })
            {
                DataReceivedEventHandler h = (s, e) => { if (e.Data != null) lock (sb) sb.AppendLine(e.Data); };
                p.OutputDataReceived += h;
                p.ErrorDataReceived += h;
                p.Start();
                lock (running) running[Thread.CurrentThread] = p;
                p.BeginOutputReadLine();
                p.BeginErrorReadLine();
                p.WaitForExit();
                lock (running) running.Remove(Thread.CurrentThread);
                lock (sb) output = sb.ToString();
                return p.ExitCode;
            }
        }

        // Chạy và ném lỗi (kèm thông báo dễ hiểu) nếu thất bại
        public static string Check(string exe, string args)
        {
            string output;
            int code = Run(exe, args, out output);
            if (code != 0)
            {
                string o = output.Trim();
                string lower = o.ToLowerInvariant();
                if (lower.Contains("invalid password") && args.Contains("--password="))
                    throw new InvalidOperationException("Sai mật khẩu.");
                if (lower.Contains("password"))
                    throw new InvalidOperationException(NeedsPasswordMessage);
                if (o.Length > 600) o = "..." + o.Substring(o.Length - 600);
                throw new InvalidOperationException(Path.GetFileNameWithoutExtension(exe) + " lỗi (mã " + code + ")" + (o.Length > 0 ? ":\r\n" + o : ""));
            }
            return output;
        }

        public const string NeedsPasswordMessage = "File PDF được bảo vệ bằng mật khẩu. Hãy dùng công cụ \"Gỡ mật khẩu\" trước.";

        public static string QpdfRun(string args) { return Check(Require(Qpdf, "qpdf"), "--warning-exit-0 " + args); }

        // Ghostscript gặp file cần mật khẩu vẫn có thể báo thành công và tạo trang trắng, nên kiểm tra trước bằng qpdf
        public static void EnsureOpenable(string pdf)
        {
            string output;
            if (Run(Require(Qpdf, "qpdf"), "--requires-password " + Q(pdf), out output) == 0)
                throw new InvalidOperationException(NeedsPasswordMessage);
        }

        public static string GsRun(string args) { return Check(Require(Gs, "Ghostscript"), "-dSAFER -dBATCH -dNOPAUSE -dQUIET " + args); }
    }

    // ------------------------------------------------------------------
    // Thông tin trang PDF (đọc qua Ghostscript)
    // ------------------------------------------------------------------
    public class PageBox
    {
        public double X0, Y0, X1, Y1;
        public int Rotate;
        public double W { get { return X1 - X0; } }
        public double H { get { return Y1 - Y0; } }
        // Kích thước khi hiển thị (đã tính xoay)
        public double VisW { get { return Rotate % 180 == 0 ? W : H; } }
        public double VisH { get { return Rotate % 180 == 0 ? H : W; } }

        // Ma trận đổi tọa độ hiển thị (u,v) sang tọa độ trang PDF: x = a*u + c*v + e, y = b*u + d*v + f
        public double[] VisualToUser()
        {
            switch (Rotate)
            {
                case 90: return new double[] { 0, 1, -1, 0, X1, Y0 };
                case 180: return new double[] { -1, 0, 0, -1, X1, Y1 };
                case 270: return new double[] { 0, -1, 1, 0, X0, Y1 };
                default: return new double[] { 1, 0, 0, 1, X0, Y0 };
            }
        }

        // Ma trận đặt ảnh (hình vuông đơn vị) vào hình chữ nhật hiển thị (u0, v0, uw, vh)
        public double[] Place(double u0, double v0, double uw, double vh)
        {
            double[] m = VisualToUser();
            double a = m[0], b = m[1], c = m[2], d = m[3], e = m[4], f = m[5];
            return new[] { a * uw, b * uw, c * vh, d * vh, a * u0 + c * v0 + e, b * u0 + d * v0 + f };
        }
    }

    public static class PdfInfo
    {
        public static List<PageBox> Pages(string file)
        {
            string gs = Ext.Require(Ext.Gs, "Ghostscript");
            Ext.EnsureOpenable(file);
            const string ps = "File (r) file runpdfbegin 1 1 pdfpagecount { pdfgetpage dup /Rotate pget not {0} if exch " +
                              "dup /TrimBox pget {exch pop} {dup /CropBox pget {exch pop} {/MediaBox pget pop} ifelse} ifelse == == } for quit";
            string output = Ext.Check(gs, "-q -dNODISPLAY -dSAFER " + Ext.Q("--permit-file-read=" + file) + " " + Ext.Q("-sFile=" + file) + " -c " + Ext.Q(ps));
            var list = new List<PageBox>();
            var lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Select(l => l.Trim()).ToList();
            for (int i = 0; i + 1 < lines.Count; i++)
            {
                if (!lines[i].StartsWith("[")) continue;
                string[] nums = lines[i].Trim('[', ']').Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (nums.Length != 4) continue;
                double[] v = nums.Select(n => double.Parse(n, CultureInfo.InvariantCulture)).ToArray();
                int rot = (int)Math.Round(double.Parse(lines[i + 1], CultureInfo.InvariantCulture));
                rot = ((rot % 360) + 360) % 360 / 90 * 90;
                list.Add(new PageBox
                {
                    X0 = Math.Min(v[0], v[2]), Y0 = Math.Min(v[1], v[3]),
                    X1 = Math.Max(v[0], v[2]), Y1 = Math.Max(v[1], v[3]),
                    Rotate = rot
                });
                i++;
            }
            if (list.Count == 0) throw new InvalidOperationException("Không đọc được trang nào trong file PDF.");
            return list;
        }
    }

    // ------------------------------------------------------------------
    // Ghi PDF từ ảnh (không cần thư viện ngoài)
    // ------------------------------------------------------------------
    public class PdfImage
    {
        public int W, H;
        public byte[] Jpeg;          // dữ liệu JPEG nhúng trực tiếp, hoặc
        public byte[] Rgb;           // điểm ảnh RGB thô (nén Flate)
        public byte[] Alpha;         // kênh trong suốt (tùy chọn)
        public float Dpi;
    }

    public class PdfPageSpec
    {
        public PageBox Box;
        public List<KeyValuePair<PdfImage, double[]>> Images = new List<KeyValuePair<PdfImage, double[]>>();
    }

    public static class PdfWriter
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        static string N(double v) { return Math.Round(v, 4).ToString("0.####", Inv); }

        public static byte[] Zlib(byte[] raw)
        {
            using (var ms = new MemoryStream())
            {
                ms.WriteByte(0x78); ms.WriteByte(0x9C);
                using (var ds = new DeflateStream(ms, CompressionMode.Compress, true)) ds.Write(raw, 0, raw.Length);
                uint a = 1, b = 0;
                foreach (byte x in raw) { a = (a + x) % 65521; b = (b + a) % 65521; }
                uint ad = (b << 16) | a;
                ms.WriteByte((byte)(ad >> 24)); ms.WriteByte((byte)(ad >> 16)); ms.WriteByte((byte)(ad >> 8)); ms.WriteByte((byte)ad);
                return ms.ToArray();
            }
        }

        public static void Save(IList<PdfPageSpec> pages, string path)
        {
            int n = pages.Count;
            int next = 3 + n;
            var imgNum = new Dictionary<PdfImage, int>();
            foreach (var pg in pages)
                foreach (var kv in pg.Images)
                    if (!imgNum.ContainsKey(kv.Key)) { imgNum[kv.Key] = next; next += kv.Key.Alpha != null ? 2 : 1; }

            var offsets = new SortedDictionary<int, long>();
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                Action<string> w = s => { byte[] bt = Encoding.ASCII.GetBytes(s); fs.Write(bt, 0, bt.Length); };
                Action<int> obj = num => { offsets[num] = fs.Position; w(num + " 0 obj\n"); };
                Action<int, string, byte[]> stream = (num, dict, data) =>
                {
                    obj(num);
                    w("<< " + dict + " /Length " + data.Length + " >>\nstream\n");
                    fs.Write(data, 0, data.Length);
                    w("\nendstream\nendobj\n");
                };

                w("%PDF-1.4\n");
                fs.Write(new byte[] { (byte)'%', 0xE2, 0xE3, 0xCF, 0xD3, (byte)'\n' }, 0, 6);
                obj(1); w("<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
                var kids = new StringBuilder();
                for (int i = 0; i < n; i++) kids.Append((3 + i) + " 0 R ");
                obj(2); w("<< /Type /Pages /Count " + n + " /Kids [" + kids + "] >>\nendobj\n");

                for (int i = 0; i < n; i++)
                {
                    PdfPageSpec pg = pages[i];
                    var content = new StringBuilder();
                    var xobj = new StringBuilder();
                    var used = new HashSet<int>();
                    foreach (var kv in pg.Images)
                    {
                        int num = imgNum[kv.Key];
                        double[] m = kv.Value;
                        content.Append("q " + string.Join(" ", m.Select(N)) + " cm /Im" + num + " Do Q\n");
                        if (used.Add(num)) xobj.Append("/Im" + num + " " + num + " 0 R ");
                    }
                    int contNo = next++;
                    PageBox b = pg.Box;
                    obj(3 + i);
                    w("<< /Type /Page /Parent 2 0 R /MediaBox [" + N(b.X0) + " " + N(b.Y0) + " " + N(b.X1) + " " + N(b.Y1) + "]" +
                      (b.Rotate != 0 ? " /Rotate " + b.Rotate : "") +
                      " /Resources << /XObject << " + xobj + ">> >> /Contents " + contNo + " 0 R >>\nendobj\n");
                    stream(contNo, "", Encoding.ASCII.GetBytes(content.ToString()));
                }

                foreach (var kv in imgNum)
                {
                    PdfImage im = kv.Key;
                    int num = kv.Value;
                    string common = "/Type /XObject /Subtype /Image /Width " + im.W + " /Height " + im.H + " /BitsPerComponent 8";
                    string smask = im.Alpha != null ? " /SMask " + (num + 1) + " 0 R" : "";
                    if (im.Jpeg != null)
                    {
                        int comps = ImageUtil.JpegComponents(im.Jpeg);
                        string cs = comps == 1 ? "/DeviceGray" : comps == 4 ? "/DeviceCMYK" : "/DeviceRGB";
                        stream(num, common + " /ColorSpace " + cs + " /Filter /DCTDecode" + (comps == 4 ? " /Decode [1 0 1 0 1 0 1 0]" : "") + smask, im.Jpeg);
                    }
                    else
                        stream(num, common + " /ColorSpace /DeviceRGB /Filter /FlateDecode" + smask, Zlib(im.Rgb));
                    if (im.Alpha != null)
                        stream(num + 1, "/Type /XObject /Subtype /Image /Width " + im.W + " /Height " + im.H +
                                        " /BitsPerComponent 8 /ColorSpace /DeviceGray /Filter /FlateDecode", Zlib(im.Alpha));
                }

                long xref = fs.Position;
                int size = next;
                w("xref\n0 " + size + "\n0000000000 65535 f \n");
                for (int i = 1; i < size; i++)
                {
                    long off;
                    w(offsets.TryGetValue(i, out off) ? off.ToString("D10") + " 00000 n \n" : "0000000000 65535 f \n");
                }
                w("trailer\n<< /Size " + size + " /Root 1 0 R >>\nstartxref\n" + xref + "\n%%EOF\n");
            }
        }
    }

    public static class ImageUtil
    {
        public static readonly string[] Extensions = { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff" };

        static ImageCodecInfo Codec(string mime)
        {
            foreach (var c in ImageCodecInfo.GetImageEncoders()) if (c.MimeType == mime) return c;
            return null;
        }

        // Đọc số kênh màu từ marker SOF của JPEG
        public static int JpegComponents(byte[] d)
        {
            int i = 2;
            while (i + 9 < d.Length)
            {
                if (d[i] != 0xFF) { i++; continue; }
                byte m = d[i + 1];
                if (m >= 0xC0 && m <= 0xCF && m != 0xC4 && m != 0xC8 && m != 0xCC) return d[i + 9];
                int len = (d[i + 2] << 8) | d[i + 3];
                i += 2 + len;
            }
            return 3;
        }

        static RotateFlipType? ExifRotation(Image img)
        {
            if (!img.PropertyIdList.Contains(0x0112)) return null;
            int o = img.GetPropertyItem(0x0112).Value[0];
            switch (o)
            {
                case 2: return RotateFlipType.RotateNoneFlipX;
                case 3: return RotateFlipType.Rotate180FlipNone;
                case 4: return RotateFlipType.Rotate180FlipX;
                case 5: return RotateFlipType.Rotate90FlipX;
                case 6: return RotateFlipType.Rotate90FlipNone;
                case 7: return RotateFlipType.Rotate270FlipX;
                case 8: return RotateFlipType.Rotate270FlipNone;
            }
            return null;
        }

        static float DpiOf(Image img)
        {
            float d = img.HorizontalResolution;
            return d >= 50 && d <= 2400 ? d : 96;
        }

        // Ảnh JPEG giữ nguyên dữ liệu gốc (không nén lại); định dạng khác lưu không mất dữ liệu.
        // TIFF nhiều trang trả về nhiều ảnh.
        public static List<PdfImage> Load(string path)
        {
            var result = new List<PdfImage>();
            byte[] bytes = File.ReadAllBytes(path);
            using (var ms = new MemoryStream(bytes))
            using (var img = Image.FromStream(ms))
            {
                bool isJpeg = bytes.Length > 3 && bytes[0] == 0xFF && bytes[1] == 0xD8;
                RotateFlipType? rot = ExifRotation(img);
                if (isJpeg && rot == null)
                {
                    result.Add(new PdfImage { W = img.Width, H = img.Height, Jpeg = bytes, Dpi = DpiOf(img) });
                    return result;
                }
                int frames = 1;
                try { frames = img.GetFrameCount(FrameDimension.Page); } catch { }
                for (int f = 0; f < frames; f++)
                {
                    if (frames > 1) img.SelectActiveFrame(FrameDimension.Page, f);
                    using (var bmp = new Bitmap(img.Width, img.Height, PixelFormat.Format24bppRgb))
                    {
                        using (var g = Graphics.FromImage(bmp))
                        {
                            g.Clear(Color.White);
                            g.DrawImage(img, new Rectangle(0, 0, img.Width, img.Height));
                        }
                        if (rot != null) bmp.RotateFlip(rot.Value);
                        var pi = isJpeg ? new PdfImage { Jpeg = ToJpeg(bmp, 92) } : new PdfImage { Rgb = RgbBytes(bmp) };
                        pi.W = bmp.Width; pi.H = bmp.Height; pi.Dpi = DpiOf(img);
                        result.Add(pi);
                    }
                }
            }
            return result;
        }

        public static byte[] ToJpeg(Image img, long quality)
        {
            var ep = new EncoderParameters(1);
            ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, quality);
            using (var ms = new MemoryStream())
            {
                img.Save(ms, Codec("image/jpeg"), ep);
                return ms.ToArray();
            }
        }

        public static byte[] RgbBytes(Bitmap bmp)
        {
            int w = bmp.Width, h = bmp.Height;
            var bd = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
            byte[] row = new byte[bd.Stride];
            byte[] rgb = new byte[w * h * 3];
            for (int y = 0; y < h; y++)
            {
                Marshal.Copy(bd.Scan0 + y * bd.Stride, row, 0, bd.Stride);
                int o = y * w * 3;
                for (int x = 0; x < w; x++)
                {
                    rgb[o + x * 3] = row[x * 3 + 2];
                    rgb[o + x * 3 + 1] = row[x * 3 + 1];
                    rgb[o + x * 3 + 2] = row[x * 3];
                }
            }
            bmp.UnlockBits(bd);
            return rgb;
        }

        // Vẽ chữ (hỗ trợ tiếng Việt) thành ảnh trong suốt để đóng dấu lên trang PDF
        public static PdfImage TextImage(string text, float sizePt, bool bold, Color color, int opacityPct, float angleDeg, int dpi)
        {
            float px = sizePt * dpi / 72f;
            using (var font = new Font("Arial", px, bold ? FontStyle.Bold : FontStyle.Regular, GraphicsUnit.Pixel))
            {
                SizeF sz;
                var fmt = StringFormat.GenericTypographic;
                using (var tmp = new Bitmap(1, 1))
                using (var g = Graphics.FromImage(tmp))
                    sz = g.MeasureString(text, font, PointF.Empty, fmt);
                double r = angleDeg * Math.PI / 180;
                double cos = Math.Abs(Math.Cos(r)), sin = Math.Abs(Math.Sin(r));
                int pad = (int)(px * 0.15) + 4;
                int w = (int)Math.Ceiling(sz.Width * cos + sz.Height * sin) + pad * 2;
                int h = (int)Math.Ceiling(sz.Width * sin + sz.Height * cos) + pad * 2;
                using (var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb))
                {
                    using (var g = Graphics.FromImage(bmp))
                    using (var br = new SolidBrush(Color.FromArgb(Math.Max(1, Math.Min(255, opacityPct * 255 / 100)), color)))
                    {
                        g.Clear(Color.Transparent);
                        g.TextRenderingHint = TextRenderingHint.AntiAlias;
                        g.SmoothingMode = SmoothingMode.AntiAlias;
                        g.TranslateTransform(w / 2f, h / 2f);
                        g.RotateTransform(-angleDeg);
                        g.DrawString(text, font, br, -sz.Width / 2f, -sz.Height / 2f, fmt);
                    }
                    var bd = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                    byte[] px32 = new byte[bd.Stride * h];
                    Marshal.Copy(bd.Scan0, px32, 0, px32.Length);
                    bmp.UnlockBits(bd);
                    byte[] alpha = new byte[w * h];
                    byte[] rgb = new byte[w * h * 3];
                    for (int y = 0; y < h; y++)
                        for (int x = 0; x < w; x++)
                        {
                            int k = y * w + x;
                            alpha[k] = px32[y * bd.Stride + x * 4 + 3];
                            rgb[k * 3] = color.R; rgb[k * 3 + 1] = color.G; rgb[k * 3 + 2] = color.B;
                        }
                    return new PdfImage { W = w, H = h, Rgb = rgb, Alpha = alpha, Dpi = dpi };
                }
            }
        }
    }

    // ------------------------------------------------------------------
    // Một lần thực hiện công cụ
    // ------------------------------------------------------------------
    public class Job
    {
        public List<string> Inputs = new List<string>();
        public string OutDir;                    // null = cùng thư mục với file gốc
        public Action<string> Log = s => { };
        public List<string> Outputs = new List<string>();
        public int Errors;
        public volatile bool Cancelled;
        readonly List<string> temps = new List<string>();

        public string Dir(string src)
        {
            string d = OutDir ?? Path.GetDirectoryName(src);
            Directory.CreateDirectory(d);
            return d;
        }

        public static string Unique(string folder, string name, string ext)
        {
            string path = Path.Combine(folder, name + ext);
            for (int k = 1; File.Exists(path) || Directory.Exists(path); k++) path = Path.Combine(folder, name + " (" + k + ")" + ext);
            return path;
        }

        public string Out(string src, string suffix, string ext)
        {
            return Unique(Dir(src), Path.GetFileNameWithoutExtension(src) + suffix, ext);
        }

        public string OutFolder(string src, string suffix)
        {
            string d = Unique(Dir(src), Path.GetFileNameWithoutExtension(src) + suffix, "");
            Directory.CreateDirectory(d);
            return d;
        }

        public string Temp()
        {
            string d = Path.Combine(Path.GetTempPath(), "pdftool_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(d);
            temps.Add(d);
            return d;
        }

        public void Cleanup()
        {
            foreach (string d in temps) try { Directory.Delete(d, true); } catch { }
            temps.Clear();
        }

        public void Done(string path)
        {
            Outputs.Add(path);
            Log("   ✔ " + path);
        }

        // Xử lý từng file; lỗi ở một file không làm dừng các file khác
        public void Each(Action<string> work)
        {
            foreach (string f in Inputs)
            {
                if (Cancelled) return;
                Log("▶ " + Path.GetFileName(f));
                try { work(f); }
                catch (Exception ex)
                {
                    if (Cancelled) return;
                    Errors++;
                    Log("   ✖ " + ex.Message);
                }
            }
        }

        public static string SizeText(long bytes)
        {
            if (bytes < 1024) return bytes + " B";
            if (bytes < 1024 * 1024) return (bytes / 1024.0).ToString("0.#") + " KB";
            return (bytes / 1024.0 / 1024.0).ToString("0.##") + " MB";
        }
    }

    // ------------------------------------------------------------------
    // Các thao tác PDF
    // ------------------------------------------------------------------
    public static class Ops
    {
        static string Q(string s) { return Ext.Q(s); }

        // "1-3, 5, 8-" -> "1-3,5,8-z" (cú pháp khoảng trang của qpdf, z = trang cuối)
        public static List<string> ParseRanges(string s)
        {
            var list = new List<string>();
            foreach (string raw in s.ToLowerInvariant().Replace(" ", "").Replace(";", ",").Split(','))
            {
                if (raw.Length == 0) continue;
                string p = raw;
                if (p.EndsWith("-")) p += "z";
                if (p.StartsWith("-")) p = "1" + p;
                if (!Regex.IsMatch(p, @"^(r?\d+|z)(-(r?\d+|z))?$") || Regex.IsMatch(p, @"(^|-)r?0+($|-)"))
                    throw new FormatException("Khoảng trang không hợp lệ: \"" + raw + "\". Ví dụ đúng: 1-3, 5, 8-z");
                list.Add(p);
            }
            if (list.Count == 0) throw new FormatException("Chưa nhập trang.");
            return list;
        }

        public static void Merge(Job job, string outName)
        {
            var sb = new StringBuilder("--empty --pages");
            foreach (string f in job.Inputs) sb.Append(" " + Q(f));
            string first = job.Inputs[0];
            string outPath = Job.Unique(job.Dir(first), outName, ".pdf");
            sb.Append(" -- " + Q(outPath));
            job.Log("▶ Gộp " + job.Inputs.Count + " file...");
            Ext.QpdfRun(sb.ToString());
            job.Done(outPath);
        }

        public static void SplitEvery(Job job, int n)
        {
            job.Each(f =>
            {
                string dir = job.OutFolder(f, "_tach");
                string pattern = Path.Combine(dir, Path.GetFileNameWithoutExtension(f) + "_%d.pdf");
                Ext.QpdfRun("--split-pages=" + n + " " + Q(f) + " " + Q(pattern));
                int count = Directory.GetFiles(dir, "*.pdf").Length;
                job.Log("   Đã tạo " + count + " file.");
                job.Done(dir);
            });
        }

        public static void SplitRanges(Job job, List<string> ranges)
        {
            job.Each(f =>
            {
                string dir = job.OutFolder(f, "_tach");
                string name = Path.GetFileNameWithoutExtension(f);
                foreach (string r in ranges)
                {
                    string outPath = Job.Unique(dir, name + "_" + r, ".pdf");
                    Ext.QpdfRun("--empty --pages " + Q(f) + " " + r + " -- " + Q(outPath));
                }
                job.Log("   Đã tạo " + ranges.Count + " file.");
                job.Done(dir);
            });
        }

        public static void SelectPages(Job job, string qpdfRange, string suffix)
        {
            job.Each(f =>
            {
                string outPath = job.Out(f, suffix, ".pdf");
                Ext.QpdfRun(Q(f) + " --pages . " + qpdfRange + " -- " + Q(outPath));
                job.Done(outPath);
            });
        }

        public static void Rotate(Job job, int angle, string range)
        {
            job.Each(f =>
            {
                string outPath = job.Out(f, "_xoay", ".pdf");
                Ext.QpdfRun(Q(f) + " " + Q(outPath) + " --rotate=+" + angle + ":" + range);
                job.Done(outPath);
            });
        }

        public static void Compress(Job job, int dpi, string preset, bool gray)
        {
            job.Each(f =>
            {
                Ext.EnsureOpenable(f);
                string outPath = job.Out(f, "_nen", ".pdf");
                string args = "-sDEVICE=pdfwrite -dCompatibilityLevel=1.5 -dPDFSETTINGS=" + preset +
                              " -dDetectDuplicateImages=true -dCompressFonts=true -dSubsetFonts=true" +
                              " -dDownsampleColorImages=true -dDownsampleGrayImages=true -dDownsampleMonoImages=true" +
                              " -dColorImageResolution=" + dpi + " -dGrayImageResolution=" + dpi + " -dMonoImageResolution=" + Math.Max(dpi * 2, 300) +
                              " -dColorImageDownsampleThreshold=1.1 -dGrayImageDownsampleThreshold=1.1" +
                              (gray ? " -sColorConversionStrategy=Gray -dProcessColorModel=/DeviceGray" : "") +
                              " " + Q("-sOutputFile=" + outPath) + " " + Q(f);
                Ext.GsRun(args);
                long before = new FileInfo(f).Length, after = new FileInfo(outPath).Length;
                if (after >= before && !gray)
                {
                    File.Delete(outPath);
                    job.Log("   File đã được tối ưu sẵn, nén không giảm được dung lượng (" + Job.SizeText(before) + "). Không tạo file mới.");
                    return;
                }
                job.Log("   " + Job.SizeText(before) + " → " + Job.SizeText(after) +
                        (after < before ? " (giảm " + (100 - after * 100 / before) + "%)" : ""));
                job.Done(outPath);
            });
        }

        public static void ImagesToPdf(Job job, bool combine, double pageW, double pageH, int orient, double marginPt)
        {
            Func<List<string>, string, int> build = (files, outPath) =>
            {
                var pages = new List<PdfPageSpec>();
                foreach (string f in files)
                {
                    if (job.Cancelled) return 0;
                    job.Log("   + " + Path.GetFileName(f));
                    foreach (PdfImage im in ImageUtil.Load(f))
                    {
                        double iw = im.W * 72.0 / im.Dpi, ih = im.H * 72.0 / im.Dpi;
                        double pw, ph;
                        if (pageW <= 0) { pw = iw; ph = ih; }
                        else
                        {
                            bool landscape = orient == 2 || (orient == 0 && im.W > im.H);
                            pw = landscape ? Math.Max(pageW, pageH) : Math.Min(pageW, pageH);
                            ph = landscape ? Math.Min(pageW, pageH) : Math.Max(pageW, pageH);
                        }
                        double m = pageW <= 0 ? 0 : marginPt;
                        double s = Math.Min((pw - 2 * m) / iw, (ph - 2 * m) / ih);
                        if (pageW > 0 && s > 1) s = 1; // không phóng to ảnh nhỏ
                        double dw = iw * s, dh = ih * s;
                        var box = new PageBox { X1 = pw, Y1 = ph };
                        var spec = new PdfPageSpec { Box = box };
                        spec.Images.Add(new KeyValuePair<PdfImage, double[]>(im, box.Place((pw - dw) / 2, (ph - dh) / 2, dw, dh)));
                        pages.Add(spec);
                    }
                }
                PdfWriter.Save(pages, outPath);
                return pages.Count;
            };

            if (combine)
            {
                string first = job.Inputs[0];
                string outPath = Job.Unique(job.Dir(first), Path.GetFileNameWithoutExtension(first), ".pdf");
                job.Log("▶ Tạo PDF từ " + job.Inputs.Count + " ảnh...");
                int n = build(job.Inputs, outPath);
                if (job.Cancelled) { try { File.Delete(outPath); } catch { } return; }
                job.Log("   " + n + " trang.");
                job.Done(outPath);
            }
            else
                job.Each(f =>
                {
                    string outPath = job.Out(f, "", ".pdf");
                    build(new List<string> { f }, outPath);
                    job.Done(outPath);
                });
        }

        public static void PdfToImages(Job job, string device, string ext, int dpi)
        {
            job.Each(f =>
            {
                Ext.EnsureOpenable(f);
                string dir = job.OutFolder(f, "_anh");
                string pattern = Path.Combine(dir, Path.GetFileNameWithoutExtension(f) + "_%03d" + ext);
                Ext.GsRun("-sDEVICE=" + device + " -r" + dpi + " -dTextAlphaBits=4 -dGraphicsAlphaBits=4 -dJPEGQ=90 " +
                          Q("-sOutputFile=" + pattern) + " " + Q(f));
                job.Log("   Đã tạo " + Directory.GetFiles(dir).Length + " ảnh.");
                job.Done(dir);
            });
        }

        public static void Protect(Job job, string userPw, string ownerPw, bool noPrint, bool noCopy, bool noModify)
        {
            if (ownerPw.Length == 0)
            {
                ownerPw = Guid.NewGuid().ToString("N");
                if (noPrint || noCopy || noModify) job.Log("Không nhập mật khẩu chủ sở hữu: dùng mật khẩu ngẫu nhiên, các hạn chế sẽ không gỡ được.");
            }
            string args = "--encrypt " + Q("--user-password=" + userPw) + " " + Q("--owner-password=" + ownerPw) + " --bits=256" +
                          (noPrint ? " --print=none" : "") + (noCopy ? " --extract=n" : "") + (noModify ? " --modify=none" : "") + " --";
            job.Each(f =>
            {
                string outPath = job.Out(f, "_baove", ".pdf");
                Ext.QpdfRun(Q(f) + " " + args + " " + Q(outPath));
                job.Done(outPath);
            });
        }

        public static void Unlock(Job job, string pw)
        {
            job.Each(f =>
            {
                string outPath = job.Out(f, "_mokhoa", ".pdf");
                Ext.QpdfRun(Q("--password=" + pw) + " --decrypt " + Q(f) + " " + Q(outPath));
                job.Done(outPath);
            });
        }

        // Đóng dấu: tạo một PDF lớp phủ cùng kích thước từng trang rồi dùng qpdf --overlay
        static void Stamp(Job job, string f, string outPath, Func<int, int, PageBox, KeyValuePair<PdfImage, double[]>?> stampFor)
        {
            List<PageBox> boxes = PdfInfo.Pages(f);
            var pages = new List<PdfPageSpec>();
            for (int i = 0; i < boxes.Count; i++)
            {
                var spec = new PdfPageSpec { Box = boxes[i] };
                var st = stampFor(i, boxes.Count, boxes[i]);
                if (st != null) spec.Images.Add(st.Value);
                pages.Add(spec);
            }
            string overlay = Path.Combine(job.Temp(), "overlay.pdf");
            PdfWriter.Save(pages, overlay);
            Ext.QpdfRun(Q(f) + " --overlay " + Q(overlay) + " --to=1-z --from=1-z -- " + Q(outPath));
        }

        public static void Watermark(Job job, string text, float sizePt, Color color, int opacity, float angle)
        {
            PdfImage img = ImageUtil.TextImage(text, sizePt, true, color, opacity, angle, 200);
            double iw = img.W * 72.0 / img.Dpi, ih = img.H * 72.0 / img.Dpi;
            job.Each(f =>
            {
                string outPath = job.Out(f, "_dongdau", ".pdf");
                Stamp(job, f, outPath, (i, count, b) =>
                {
                    // Thu nhỏ nếu dấu lớn hơn trang
                    double s = Math.Min(1, Math.Min(b.VisW * 0.95 / iw, b.VisH * 0.95 / ih));
                    double w = iw * s, h = ih * s;
                    return new KeyValuePair<PdfImage, double[]>(img, b.Place((b.VisW - w) / 2, (b.VisH - h) / 2, w, h));
                });
                job.Done(outPath);
            });
        }

        // position: 0 dưới giữa, 1 dưới phải, 2 dưới trái, 3 trên giữa, 4 trên phải, 5 trên trái
        public static void PageNumbers(Job job, string format, int position, int fromPage, int startNumber, float sizePt, double marginPt)
        {
            job.Each(f =>
            {
                string outPath = job.Out(f, "_sotrang", ".pdf");
                Stamp(job, f, outPath, (i, count, b) =>
                {
                    if (i + 1 < fromPage) return null;
                    int total = count - fromPage + startNumber;
                    int num = i + 1 - fromPage + startNumber;
                    string text = format.Replace("{n}", num.ToString()).Replace("{N}", total.ToString());
                    PdfImage img = ImageUtil.TextImage(text, sizePt, false, Color.Black, 100, 0, 300);
                    double w = img.W * 72.0 / img.Dpi, h = img.H * 72.0 / img.Dpi;
                    double u = position % 3 == 0 ? (b.VisW - w) / 2 : position % 3 == 1 ? b.VisW - marginPt - w : marginPt;
                    double v = position < 3 ? marginPt - h * 0.25 : b.VisH - marginPt - h * 0.75;
                    return new KeyValuePair<PdfImage, double[]>(img, b.Place(u, v, w, h));
                });
                job.Done(outPath);
            });
        }

        // ---------------- OCR ----------------
        public static string TessDataDir
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PdfTool", "tessdata"); }
        }

        // Trả về thư mục tessdata có đủ ngôn ngữ; tải dữ liệu còn thiếu từ GitHub (tesseract-ocr/tessdata_fast)
        static string EnsureTessData(Job job, string[] langs)
        {
            string tess = Ext.Require(Ext.Tesseract, "Tesseract OCR");
            string builtIn = Path.Combine(Path.GetDirectoryName(tess), "tessdata");
            if (langs.All(l => File.Exists(Path.Combine(builtIn, l + ".traineddata")))) return builtIn;

            string dir = TessDataDir;
            Directory.CreateDirectory(Path.Combine(dir, "configs"));
            if (Directory.Exists(Path.Combine(builtIn, "configs")))
                foreach (string c in Directory.GetFiles(Path.Combine(builtIn, "configs")))
                {
                    string dst = Path.Combine(dir, "configs", Path.GetFileName(c));
                    if (!File.Exists(dst)) File.Copy(c, dst);
                }
            if (!File.Exists(Path.Combine(dir, "pdf.ttf")) && File.Exists(Path.Combine(builtIn, "pdf.ttf")))
                File.Copy(Path.Combine(builtIn, "pdf.ttf"), Path.Combine(dir, "pdf.ttf"));
            if (!File.Exists(Path.Combine(dir, "configs", "pdf")))
                File.WriteAllText(Path.Combine(dir, "configs", "pdf"), "tessedit_create_pdf 1\n");

            foreach (string l in langs)
            {
                string dst = Path.Combine(dir, l + ".traineddata");
                if (File.Exists(dst)) continue;
                job.Log("   Đang tải dữ liệu nhận dạng ngôn ngữ \"" + l + "\" (chỉ lần đầu)...");
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
                string tmp = dst + ".part";
                using (var wc = new WebClient())
                    wc.DownloadFile("https://github.com/tesseract-ocr/tessdata_fast/raw/main/" + l + ".traineddata", tmp);
                File.Move(tmp, dst);
            }
            return dir;
        }

        public static void Ocr(Job job, string[] langs, int dpi)
        {
            string tess = Ext.Require(Ext.Tesseract, "Tesseract OCR");
            job.Log("Chuẩn bị dữ liệu nhận dạng...");
            string tessData = EnsureTessData(job, langs);
            job.Each(f =>
            {
                Ext.EnsureOpenable(f);
                string outPath = job.Out(f, "_ocr", ".pdf");
                string tmp = job.Temp();
                job.Log("   Đang chuyển trang thành ảnh...");
                Ext.GsRun("-sDEVICE=pnggray -r" + dpi + " " + Q("-sOutputFile=" + Path.Combine(tmp, "p-%05d.png")) + " " + Q(f));
                string[] imgs = Directory.GetFiles(tmp, "p-*.png").OrderBy(x => x).ToArray();
                File.WriteAllLines(Path.Combine(tmp, "list.txt"), imgs, new UTF8Encoding(false));
                job.Log("   Đang nhận dạng chữ " + imgs.Length + " trang (có thể mất vài phút)...");
                // Chỉ tạo lớp chữ ẩn rồi phủ lên file gốc: giữ nguyên chất lượng và dung lượng file gốc
                Ext.Check(tess, Q(Path.Combine(tmp, "list.txt")) + " " + Q(Path.Combine(tmp, "ocr")) +
                               " --tessdata-dir " + Q(tessData) + " -l " + string.Join("+", langs) +
                               " --dpi " + dpi + " -c textonly_pdf=1 pdf");
                string textPdf = Path.Combine(tmp, "ocr.pdf");
                if (!File.Exists(textPdf)) throw new InvalidOperationException("Tesseract không tạo được kết quả.");
                Ext.QpdfRun(Q(f) + " --overlay " + Q(textPdf) + " --to=1-z --from=1-z -- " + Q(outPath));
                job.Done(outPath);
            });
        }
    }

    // ------------------------------------------------------------------
    // Định nghĩa công cụ và giao diện tùy chọn
    // ------------------------------------------------------------------
    class Opts : TableLayoutPanel
    {
        readonly float scale;
        int S(int v) { return (int)(v * scale); }

        public Opts(float scale)
        {
            this.scale = scale;
            ColumnCount = 3;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Dock = DockStyle.Top;
            ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        }

        T Row<T>(string label, T c, string hint) where T : Control
        {
            int r = RowCount++;
            RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(new Label { Text = label, AutoSize = true, Margin = new Padding(S(3), S(7), S(10), S(3)) }, 0, r);
            c.Margin = new Padding(S(3));
            Controls.Add(c, 1, r);
            if (hint != null) Controls.Add(new Label { Text = hint, AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(S(6), S(7), S(3), S(3)) }, 2, r);
            return c;
        }

        public ComboBox Combo(string label, int width, params string[] items)
        {
            var c = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = S(width) };
            c.Items.AddRange(items);
            c.SelectedIndex = 0;
            return Row(label, c, null);
        }

        public TextBox Input(string label, int width, string value, string hint, bool password = false)
        {
            return Row(label, new TextBox { Width = S(width), Text = value, UseSystemPasswordChar = password }, hint);
        }

        public NumericUpDown Num(string label, int min, int max, int value, string hint)
        {
            return Row(label, new NumericUpDown { Minimum = min, Maximum = max, Value = value, Width = S(80) }, hint);
        }

        public CheckBox Check(string text, bool value)
        {
            int r = RowCount++;
            RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var c = new CheckBox { Text = text, Checked = value, AutoSize = true, Margin = new Padding(S(3), S(5), S(3), S(3)) };
            Controls.Add(c, 1, r);
            SetColumnSpan(c, 2);
            return c;
        }

        public Button ColorPick(string label, Color value)
        {
            var b = new Button { Width = S(80), Height = S(26), BackColor = value, FlatStyle = FlatStyle.Flat };
            b.Click += (s, e) =>
            {
                using (var d = new ColorDialog { Color = b.BackColor, FullOpen = true })
                    if (d.ShowDialog(FindForm()) == DialogResult.OK) b.BackColor = d.Color;
            };
            return Row(label, b, null);
        }

        public void Note(string text)
        {
            int r = RowCount++;
            RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var l = new Label { Text = text, AutoSize = true, ForeColor = Color.DimGray, MaximumSize = new Size(S(620), 0), Margin = new Padding(S(3), S(6), S(3), S(3)) };
            Controls.Add(l, 0, r);
            SetColumnSpan(l, 3);
        }
    }

    class ToolDef
    {
        public string Name, Short, Desc;
        public Color Color;
        public bool ImageInput;              // đầu vào là ảnh (ngược lại: PDF)
        public int MinFiles = 1;
        public Opts Options;
        public Func<Action<Job>> Prepare;    // gọi trên luồng giao diện: đọc tùy chọn, trả về công việc chạy nền (null = hủy)
        public readonly List<string> Files = new List<string>();
        public override string ToString() { return Name; }
    }

    static class ToolCatalog
    {
        static readonly double Mm = 72 / 25.4;

        static void Warn(string msg) { MessageBox.Show(msg, "PdfTool", MessageBoxButtons.OK, MessageBoxIcon.Warning); }

        // Kiểm tra cú pháp khoảng trang ngay trên giao diện
        static List<string> Ranges(string s)
        {
            try { return Ops.ParseRanges(s); }
            catch (FormatException ex) { Warn(ex.Message); return null; }
        }

        public static List<ToolDef> Build(float scale)
        {
            var tools = new List<ToolDef>();
            const string rangeHint = "VD: 1-3, 5, 8-z  (z = trang cuối)";

            // --- Gộp PDF
            {
                var t = new ToolDef { Name = "Gộp PDF", Short = "GỘP", Color = Color.FromArgb(0, 120, 212), MinFiles = 2,
                    Desc = "Nối nhiều file PDF thành một file theo thứ tự trong danh sách (dùng nút Lên/Xuống để sắp xếp)." };
                var o = t.Options = new Opts(scale);
                var name = o.Input("Tên file kết quả", 260, "Gop_PDF", ".pdf");
                t.Prepare = () =>
                {
                    string n = name.Text.Trim();
                    foreach (char c in Path.GetInvalidFileNameChars()) n = n.Replace(c, '_');
                    if (n.Length == 0) n = "Gop_PDF";
                    return job => Ops.Merge(job, n);
                };
                tools.Add(t);
            }

            // --- Tách PDF
            {
                var t = new ToolDef { Name = "Tách PDF", Short = "TÁCH", Color = Color.FromArgb(16, 137, 62),
                    Desc = "Tách một file PDF thành nhiều file: mỗi trang một file, mỗi N trang một file, hoặc theo các khoảng trang." };
                var o = t.Options = new Opts(scale);
                var mode = o.Combo("Cách tách", 260, "Mỗi trang thành 1 file", "Mỗi N trang thành 1 file", "Theo khoảng trang (mỗi khoảng 1 file)");
                var n = o.Num("Số trang mỗi file (N)", 1, 10000, 2, null);
                var ranges = o.Input("Các khoảng trang", 260, "1-3, 4-z", "mỗi khoảng cách nhau dấu phẩy");
                EventHandler sync = (s, e) => { n.Enabled = mode.SelectedIndex == 1; ranges.Enabled = mode.SelectedIndex == 2; };
                mode.SelectedIndexChanged += sync; sync(null, null);
                t.Prepare = () =>
                {
                    if (mode.SelectedIndex == 2)
                    {
                        var r = Ranges(ranges.Text);
                        if (r == null) return null;
                        return job => Ops.SplitRanges(job, r);
                    }
                    int every = mode.SelectedIndex == 0 ? 1 : (int)n.Value;
                    return job => Ops.SplitEvery(job, every);
                };
                tools.Add(t);
            }

            // --- Trích xuất / sắp xếp trang
            {
                var t = new ToolDef { Name = "Trích xuất / sắp xếp trang", Short = "TRÍCH", Color = Color.FromArgb(0, 153, 188),
                    Desc = "Tạo file mới chỉ gồm các trang được chọn, theo đúng thứ tự nhập. Dùng để lấy vài trang, sắp xếp lại hoặc đảo ngược thứ tự (z-1)." };
                var o = t.Options = new Opts(scale);
                var pages = o.Input("Trang cần lấy", 260, "1-3", rangeHint);
                o.Note("Ví dụ: \"3, 1, 2\" đổi thứ tự 3 trang đầu;  \"z-1\" đảo ngược toàn bộ;  \"1-z, 1-z\" nhân đôi tài liệu.");
                t.Prepare = () =>
                {
                    var r = Ranges(pages.Text);
                    if (r == null) return null;
                    return job => Ops.SelectPages(job, string.Join(",", r), "_trich");
                };
                tools.Add(t);
            }

            // --- Xóa trang
            {
                var t = new ToolDef { Name = "Xóa trang", Short = "XÓA", Color = Color.FromArgb(209, 52, 56),
                    Desc = "Xóa các trang không cần thiết khỏi file PDF." };
                var o = t.Options = new Opts(scale);
                var pages = o.Input("Trang cần xóa", 260, "1", rangeHint);
                t.Prepare = () =>
                {
                    var r = Ranges(pages.Text);
                    if (r == null) return null;
                    return job => Ops.SelectPages(job, "1-z," + string.Join(",", r.Select(x => "x" + x)), "_daxoa");
                };
                tools.Add(t);
            }

            // --- Xoay trang
            {
                var t = new ToolDef { Name = "Xoay trang", Short = "XOAY", Color = Color.FromArgb(135, 100, 184),
                    Desc = "Xoay tất cả hoặc một số trang của file PDF." };
                var o = t.Options = new Opts(scale);
                var angle = o.Combo("Góc xoay", 200, "90° theo chiều kim đồng hồ", "180°", "90° ngược chiều kim đồng hồ");
                var pages = o.Input("Trang cần xoay", 200, "1-z", rangeHint);
                t.Prepare = () =>
                {
                    var r = Ranges(pages.Text);
                    if (r == null) return null;
                    int a = new[] { 90, 180, 270 }[angle.SelectedIndex];
                    return job => Ops.Rotate(job, a, string.Join(",", r));
                };
                tools.Add(t);
            }

            // --- Nén PDF
            {
                var t = new ToolDef { Name = "Nén PDF", Short = "NÉN", Color = Color.FromArgb(202, 80, 16),
                    Desc = "Giảm dung lượng file PDF bằng cách giảm độ phân giải ảnh và tối ưu phông chữ." };
                var o = t.Options = new Opts(scale);
                var level = o.Combo("Mức nén", 300, "Khuyến nghị - 150 dpi (cân bằng)", "Nén mạnh - 96 dpi (file nhỏ nhất)", "Nén nhẹ - 200 dpi (giữ chất lượng in)", "Chất lượng cao - 300 dpi");
                var gray = o.Check("Chuyển sang đen trắng (thang xám) để giảm thêm dung lượng", false);
                t.Prepare = () =>
                {
                    int dpi = new[] { 150, 96, 200, 300 }[level.SelectedIndex];
                    string preset = new[] { "/ebook", "/screen", "/printer", "/prepress" }[level.SelectedIndex];
                    bool g = gray.Checked;
                    return job => Ops.Compress(job, dpi, preset, g);
                };
                tools.Add(t);
            }

            // --- Ảnh sang PDF
            {
                var t = new ToolDef { Name = "Ảnh sang PDF", Short = "ẢNH→", Color = Color.FromArgb(227, 0, 140), ImageInput = true,
                    Desc = "Chuyển ảnh JPG, PNG, BMP, GIF, TIFF thành PDF. Ảnh JPG được nhúng nguyên gốc, không giảm chất lượng." };
                var o = t.Options = new Opts(scale);
                var size = o.Combo("Khổ trang", 220, "A4", "Theo kích thước ảnh", "A3", "A5", "Letter");
                var orient = o.Combo("Hướng trang", 220, "Tự động theo ảnh", "Dọc", "Ngang");
                var margin = o.Num("Lề (mm)", 0, 50, 10, null);
                var combine = o.Check("Gộp tất cả ảnh vào một file PDF", true);
                size.SelectedIndexChanged += (s, e) => orient.Enabled = margin.Enabled = size.SelectedIndex != 1;
                t.Prepare = () =>
                {
                    double[][] sizes = { new[] { 210.0, 297 }, new[] { 0.0, 0 }, new[] { 297.0, 420 }, new[] { 148.0, 210 }, new[] { 215.9, 279.4 } };
                    double[] sz = sizes[size.SelectedIndex];
                    double w = sz[0] * Mm, h = sz[1] * Mm, m = (double)margin.Value * Mm;
                    int or = orient.SelectedIndex;
                    bool c = combine.Checked;
                    return job => Ops.ImagesToPdf(job, c, w, h, or, m);
                };
                tools.Add(t);
            }

            // --- PDF sang ảnh
            {
                var t = new ToolDef { Name = "PDF sang ảnh", Short = "→ẢNH", Color = Color.FromArgb(0, 99, 177),
                    Desc = "Xuất mỗi trang PDF thành một file ảnh PNG hoặc JPG." };
                var o = t.Options = new Opts(scale);
                var fmt = o.Combo("Định dạng", 200, "PNG (sắc nét)", "JPG (file nhỏ)");
                var color = o.Combo("Màu", 200, "Màu", "Thang xám");
                var dpi = o.Combo("Độ phân giải", 200, "150 dpi", "72 dpi", "100 dpi", "200 dpi", "300 dpi", "600 dpi");
                t.Prepare = () =>
                {
                    bool png = fmt.SelectedIndex == 0, g = color.SelectedIndex == 1;
                    string dev = png ? (g ? "pnggray" : "png16m") : (g ? "jpeggray" : "jpeg");
                    int d = int.Parse(dpi.Text.Split(' ')[0]);
                    return job => Ops.PdfToImages(job, dev, png ? ".png" : ".jpg", d);
                };
                tools.Add(t);
            }

            // --- OCR
            {
                var t = new ToolDef { Name = "Nhận dạng chữ (OCR)", Short = "OCR", Color = Color.FromArgb(0, 128, 128),
                    Desc = "Nhận dạng chữ trong PDF scan để có thể tìm kiếm và sao chép văn bản. Giữ nguyên hình ảnh gốc, chỉ thêm lớp chữ ẩn." };
                var o = t.Options = new Opts(scale);
                var lang = o.Combo("Ngôn ngữ", 220, "Tiếng Việt + Tiếng Anh", "Tiếng Việt", "Tiếng Anh");
                var dpi = o.Combo("Độ phân giải quét", 220, "300 dpi (khuyến nghị)", "400 dpi (chữ nhỏ)", "200 dpi (nhanh)");
                o.Note("Lần đầu sử dụng sẽ tải dữ liệu ngôn ngữ (~5 MB) từ github.com/tesseract-ocr về " + Ops.TessDataDir + ".");
                t.Prepare = () =>
                {
                    string[] langs = new[] { new[] { "vie", "eng" }, new[] { "vie" }, new[] { "eng" } }[lang.SelectedIndex];
                    int d = new[] { 300, 400, 200 }[dpi.SelectedIndex];
                    return job => Ops.Ocr(job, langs, d);
                };
                tools.Add(t);
            }

            // --- Đánh số trang
            {
                var t = new ToolDef { Name = "Đánh số trang", Short = "SỐ", Color = Color.FromArgb(73, 130, 5),
                    Desc = "Thêm số trang vào đầu hoặc cuối mỗi trang." };
                var o = t.Options = new Opts(scale);
                var pos = o.Combo("Vị trí", 220, "Dưới - giữa", "Dưới - phải", "Dưới - trái", "Trên - giữa", "Trên - phải", "Trên - trái");
                var fmt = o.Combo("Kiểu", 220, "1", "1 / 10", "Trang 1", "Trang 1 / 10", "- 1 -");
                var size = o.Num("Cỡ chữ (pt)", 6, 48, 11, null);
                var from = o.Num("Bắt đầu đánh từ trang", 1, 100000, 1, "VD: 2 để bỏ qua trang bìa");
                var start = o.Num("Số đầu tiên", 0, 100000, 1, null);
                var margin = o.Num("Khoảng cách mép (mm)", 3, 50, 10, null);
                t.Prepare = () =>
                {
                    string f = new[] { "{n}", "{n} / {N}", "Trang {n}", "Trang {n} / {N}", "- {n} -" }[fmt.SelectedIndex];
                    int p = pos.SelectedIndex, fr = (int)from.Value, st = (int)start.Value;
                    float sz = (float)size.Value;
                    double m = (double)margin.Value * Mm;
                    return job => Ops.PageNumbers(job, f, p, fr, st, sz, m);
                };
                tools.Add(t);
            }

            // --- Đóng dấu chữ (watermark)
            {
                var t = new ToolDef { Name = "Đóng dấu chữ (watermark)", Short = "DẤU", Color = Color.FromArgb(194, 57, 179),
                    Desc = "Đóng dấu chữ mờ lên giữa tất cả các trang, ví dụ \"BẢN SAO\", \"MẬT\", tên công ty." };
                var o = t.Options = new Opts(scale);
                var text = o.Input("Nội dung", 300, "BẢN SAO", null);
                var size = o.Num("Cỡ chữ (pt)", 8, 300, 72, null);
                var color = o.ColorPick("Màu", Color.FromArgb(200, 0, 0));
                var opacity = o.Num("Độ đậm (%)", 5, 100, 25, null);
                var angle = o.Combo("Hướng", 200, "Chéo 45°", "Ngang", "Dọc");
                t.Prepare = () =>
                {
                    string s = text.Text.Trim();
                    if (s.Length == 0) { Warn("Hãy nhập nội dung dấu."); return null; }
                    float sz = (float)size.Value, a = new[] { 45f, 0f, 90f }[angle.SelectedIndex];
                    Color c = color.BackColor;
                    int op = (int)opacity.Value;
                    return job => Ops.Watermark(job, s, sz, c, op, a);
                };
                tools.Add(t);
            }

            // --- Đặt mật khẩu
            {
                var t = new ToolDef { Name = "Đặt mật khẩu", Short = "KHÓA", Color = Color.FromArgb(118, 118, 118),
                    Desc = "Mã hóa PDF bằng AES-256: cần mật khẩu để mở, và/hoặc hạn chế in, sao chép, chỉnh sửa." };
                var o = t.Options = new Opts(scale);
                var user = o.Input("Mật khẩu mở file", 220, "", "để trống nếu chỉ muốn hạn chế quyền", true);
                var user2 = o.Input("Nhập lại mật khẩu", 220, "", null, true);
                var owner = o.Input("Mật khẩu chủ sở hữu", 220, "", "dùng để gỡ hạn chế (tùy chọn)", true);
                var noPrint = o.Check("Không cho in", false);
                var noCopy = o.Check("Không cho sao chép nội dung", false);
                var noModify = o.Check("Không cho chỉnh sửa", false);
                t.Prepare = () =>
                {
                    if (user.Text != user2.Text) { Warn("Hai lần nhập mật khẩu không khớp."); return null; }
                    bool restrict = noPrint.Checked || noCopy.Checked || noModify.Checked;
                    if (user.Text.Length == 0 && !restrict) { Warn("Hãy nhập mật khẩu mở file hoặc chọn ít nhất một hạn chế."); return null; }
                    string u = user.Text, ow = owner.Text;
                    bool p = noPrint.Checked, c = noCopy.Checked, m = noModify.Checked;
                    return job => Ops.Protect(job, u, ow, p, c, m);
                };
                tools.Add(t);
            }

            // --- Gỡ mật khẩu
            {
                var t = new ToolDef { Name = "Gỡ mật khẩu", Short = "MỞ", Color = Color.FromArgb(96, 94, 92),
                    Desc = "Tạo bản PDF không còn mật khẩu và hạn chế (cần biết mật khẩu nếu file yêu cầu mật khẩu khi mở)." };
                var o = t.Options = new Opts(scale);
                var pw = o.Input("Mật khẩu", 220, "", "để trống nếu file mở được mà không cần mật khẩu", true);
                t.Prepare = () =>
                {
                    string p = pw.Text;
                    return job => Ops.Unlock(job, p);
                };
                tools.Add(t);
            }

            return tools;
        }
    }

    // ------------------------------------------------------------------
    // Thẻ "Hộp công cụ"
    // ------------------------------------------------------------------
    class ToolboxView : UserControl
    {
        readonly List<ToolDef> tools;
        ToolDef tool;
        Job running;
        Thread worker;
        string lastResult;
        float scale = 1f;

        ListBox lstTools;
        Label lblTitle, lblDesc;
        ListView lvFiles;
        Panel pnlOptions;
        RadioButton rbSame, rbOther;
        TextBox txtOut;
        Button btnAdd, btnRemove, btnUp, btnDown, btnClear, btnBrowse, btnRun, btnCancel, btnOpen;
        TextBox txtLog;
        ProgressBar progress;

        int S(int v) { return (int)(v * scale); }

        public ToolboxView(float scale)
        {
            this.scale = scale;
            Font = new Font("Segoe UI", 9f);
            AllowDrop = true;
            DragEnter += (s, e) => e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) && running == null ? DragDropEffects.Copy : DragDropEffects.None;
            DragDrop += (s, e) => AddFiles((string[])e.Data.GetData(DataFormats.FileDrop));

            tools = ToolCatalog.Build(scale);
            BuildUi();
            LoadSettings();
            CheckTools();
        }

        public bool IsBusy { get { return running != null; } }

        public void CancelRunning()
        {
            if (running == null) return;
            running.Cancelled = true;
            Ext.Kill(worker);
        }

        Button MakeButton(string text, EventHandler onClick, int width)
        {
            var b = new Button { Text = text, Width = S(width), Height = S(30), Margin = new Padding(S(2)) };
            b.Click += onClick;
            return b;
        }

        GroupBox Group(string text)
        {
            return new GroupBox { Text = text, Dock = DockStyle.Fill, Padding = new Padding(S(8), S(4), S(8), S(6)), Margin = new Padding(S(4)) };
        }

        void BuildUi()
        {
            // --- Danh sách công cụ bên trái
            lstTools = new ListBox
            {
                Dock = DockStyle.Left, Width = S(250), DrawMode = DrawMode.OwnerDrawFixed, ItemHeight = S(46),
                BorderStyle = BorderStyle.None, BackColor = Color.FromArgb(243, 243, 243), IntegralHeight = false
            };
            foreach (var t in tools) lstTools.Items.Add(t);
            lstTools.DrawItem += DrawTool;
            lstTools.SelectedIndexChanged += (s, e) => ShowTool(lstTools.SelectedItem as ToolDef);

            // --- Khu vực bên phải
            var right = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, Padding = new Padding(S(10), S(6), S(10), S(6)) };
            right.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // tiêu đề
            right.RowStyles.Add(new RowStyle(SizeType.Percent, 55));    // danh sách file
            right.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // tùy chọn
            right.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // thư mục lưu
            right.RowStyles.Add(new RowStyle(SizeType.AutoSize));       // nút thực hiện
            right.RowStyles.Add(new RowStyle(SizeType.Percent, 45));    // nhật ký

            var head = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false, Dock = DockStyle.Fill, Margin = new Padding(S(4), 0, S(4), S(4)) };
            lblTitle = new Label { AutoSize = true, Font = new Font("Segoe UI Semibold", 16f), Margin = new Padding(0) };
            lblDesc = new Label { AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(S(2), S(2), 0, S(2)) };
            head.Controls.Add(lblTitle);
            head.Controls.Add(lblDesc);
            right.Resize += (s, e) => lblDesc.MaximumSize = new Size(Math.Max(S(200), right.ClientSize.Width - S(40)), 0);

            // Danh sách file
            var gFiles = Group("File cần xử lý (kéo thả file vào cửa sổ)");
            lvFiles = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, HideSelection = false, AllowDrop = true };
            lvFiles.Columns.Add("Tên file", S(330));
            lvFiles.Columns.Add("Kích thước", S(90), HorizontalAlignment.Right);
            lvFiles.Columns.Add("Thư mục", S(330));
            lvFiles.DragEnter += (s, e) => e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) && running == null ? DragDropEffects.Copy : DragDropEffects.None;
            lvFiles.DragDrop += (s, e) => AddFiles((string[])e.Data.GetData(DataFormats.FileDrop));
            lvFiles.SelectedIndexChanged += (s, e) => UpdateButtons();
            lvFiles.KeyDown += (s, e) => { if (e.KeyCode == Keys.Delete) RemoveSelected(); };
            var fileBtns = new FlowLayoutPanel { Dock = DockStyle.Right, FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false, Padding = new Padding(S(6), 0, 0, 0) };
            btnAdd = MakeButton("Thêm file...", (s, e) => BrowseFiles(), 130);
            btnRemove = MakeButton("Bỏ khỏi danh sách", (s, e) => RemoveSelected(), 130);
            btnUp = MakeButton("▲ Lên", (s, e) => MoveSelected(-1), 130);
            btnDown = MakeButton("▼ Xuống", (s, e) => MoveSelected(1), 130);
            btnClear = MakeButton("Bỏ tất cả", (s, e) => { tool.Files.Clear(); RefreshFiles(); }, 130);
            fileBtns.Controls.AddRange(new Control[] { btnAdd, btnRemove, btnUp, btnDown, btnClear });
            gFiles.Controls.Add(lvFiles);
            gFiles.Controls.Add(fileBtns);

            // Tùy chọn
            var gOpt = Group("Tùy chọn");
            gOpt.AutoSize = true;
            pnlOptions = new Panel { Dock = DockStyle.Fill, AutoSize = true };
            gOpt.Controls.Add(pnlOptions);

            // Thư mục lưu
            var gOut = Group("Lưu kết quả");
            gOut.AutoSize = true;
            var outFlow = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true };
            rbSame = new RadioButton { Text = "Cùng thư mục với file gốc", AutoSize = true, Checked = true, Margin = new Padding(S(3), S(7), S(12), S(3)) };
            rbOther = new RadioButton { Text = "Thư mục:", AutoSize = true, Margin = new Padding(S(3), S(7), S(3), S(3)) };
            txtOut = new TextBox { Width = S(380), Margin = new Padding(S(3), S(5), S(3), S(3)) };
            btnBrowse = MakeButton("...", (s, e) => BrowseOut(), 34);
            btnBrowse.Height = S(26);
            txtOut.TextChanged += (s, e) => { if (txtOut.Focused) rbOther.Checked = true; };
            outFlow.Controls.AddRange(new Control[] { rbSame, rbOther, txtOut, btnBrowse });
            gOut.Controls.Add(outFlow);

            // Nút thực hiện
            var actions = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = false, Margin = new Padding(S(4), S(6), S(4), S(2)) };
            btnRun = MakeButton("THỰC HIỆN", (s, e) => RunTool(), 200);
            btnRun.Height = S(40);
            btnRun.Font = new Font(Font.FontFamily, 11f, FontStyle.Bold);
            btnRun.BackColor = Color.FromArgb(0, 120, 212);
            btnRun.ForeColor = Color.White;
            btnRun.FlatStyle = FlatStyle.Flat;
            btnCancel = MakeButton("Dừng", (s, e) => CancelRunning(), 90);
            btnCancel.Height = S(40);
            btnOpen = MakeButton("Mở thư mục kết quả", (s, e) => OpenResult(), 160);
            btnOpen.Height = S(40);
            progress = new ProgressBar { Style = ProgressBarStyle.Marquee, Width = S(160), Height = S(18), Visible = false, Margin = new Padding(S(12), S(13), S(3), S(3)) };
            actions.Controls.AddRange(new Control[] { btnRun, btnCancel, btnOpen, progress });

            // Nhật ký
            txtLog = new TextBox
            {
                Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
                BackColor = Color.White, Font = new Font("Consolas", 9f), Margin = new Padding(S(4))
            };

            right.Controls.Add(head, 0, 0);
            right.Controls.Add(gFiles, 0, 1);
            right.Controls.Add(gOpt, 0, 2);
            right.Controls.Add(gOut, 0, 3);
            right.Controls.Add(actions, 0, 4);
            right.Controls.Add(txtLog, 0, 5);

            Controls.Add(right);
            Controls.Add(new Panel { Dock = DockStyle.Left, Width = 1, BackColor = Color.FromArgb(220, 220, 220) });
            Controls.Add(lstTools);
            lstTools.SelectedIndex = 0;
        }

        void DrawTool(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            var t = (ToolDef)lstTools.Items[e.Index];
            bool sel = (e.State & DrawItemState.Selected) != 0;
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            using (var bg = new SolidBrush(sel ? Color.FromArgb(204, 228, 247) : lstTools.BackColor)) g.FillRectangle(bg, e.Bounds);
            if (sel) using (var bar = new SolidBrush(Color.FromArgb(0, 120, 212))) g.FillRectangle(bar, e.Bounds.X, e.Bounds.Y + S(6), S(3), e.Bounds.Height - S(12));

            var icon = new Rectangle(e.Bounds.X + S(12), e.Bounds.Y + S(7), S(32), S(32));
            using (var path = RoundRect(icon, S(6)))
            using (var b = new SolidBrush(t.Color)) g.FillPath(b, path);
            using (var f = new Font("Segoe UI", t.Short.Length > 3 ? 6.5f : 7.5f, FontStyle.Bold))
            using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                g.DrawString(t.Short, f, Brushes.White, icon, sf);
            var textRect = new Rectangle(icon.Right + S(10), e.Bounds.Y, e.Bounds.Width - icon.Right - S(12), e.Bounds.Height);
            TextRenderer.DrawText(g, t.Name, Font, textRect, Color.FromArgb(32, 32, 32), TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis);
        }

        static GraphicsPath RoundRect(Rectangle r, int rad)
        {
            var p = new GraphicsPath();
            int d = rad * 2;
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        // ------------------------------------------------------------------
        void ShowTool(ToolDef t)
        {
            if (t == null) return;
            tool = t;
            lblTitle.Text = t.Name;
            lblDesc.Text = t.Desc;
            pnlOptions.Controls.Clear();
            pnlOptions.Controls.Add(t.Options);
            lvFiles.Parent.Text = (t.ImageInput ? "Ảnh cần chuyển" : t.MinFiles > 1 ? "Các file PDF cần gộp" : "File PDF cần xử lý") + " (kéo thả file vào cửa sổ)";
            RefreshFiles();
        }

        bool Accepts(string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            return tool.ImageInput ? ImageUtil.Extensions.Contains(ext) : ext == ".pdf";
        }

        public void AddFiles(IEnumerable<string> paths)
        {
            if (running != null) return;
            int skipped = 0;
            foreach (string p in paths)
            {
                IEnumerable<string> files = Directory.Exists(p) ? Directory.GetFiles(p).OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase) : (IEnumerable<string>)new[] { p };
                foreach (string f in files)
                {
                    if (!File.Exists(f) || !Accepts(f)) { skipped++; continue; }
                    if (!tool.Files.Contains(f, StringComparer.OrdinalIgnoreCase)) tool.Files.Add(f);
                }
            }
            RefreshFiles();
            if (skipped > 0)
                Log("Bỏ qua " + skipped + " file không phải " + (tool.ImageInput ? "ảnh" : "PDF") + ".");
        }

        void BrowseFiles()
        {
            using (var d = new OpenFileDialog { Multiselect = true })
            {
                d.Filter = tool.ImageInput
                    ? "Ảnh (*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff)|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff|Tất cả file|*.*"
                    : "File PDF (*.pdf)|*.pdf|Tất cả file|*.*";
                if (d.ShowDialog(this) == DialogResult.OK) AddFiles(d.FileNames);
            }
        }

        void RefreshFiles()
        {
            var sel = lvFiles.SelectedIndices.Cast<int>().Select(i => tool.Files.ElementAtOrDefault(i)).ToList();
            lvFiles.BeginUpdate();
            lvFiles.Items.Clear();
            foreach (string f in tool.Files)
            {
                string size = "";
                try { size = Job.SizeText(new FileInfo(f).Length); } catch { }
                var it = new ListViewItem(new[] { Path.GetFileName(f), size, Path.GetDirectoryName(f) });
                it.Selected = sel.Contains(f);
                lvFiles.Items.Add(it);
            }
            lvFiles.EndUpdate();
            UpdateButtons();
        }

        void RemoveSelected()
        {
            foreach (int i in lvFiles.SelectedIndices.Cast<int>().OrderByDescending(i => i)) tool.Files.RemoveAt(i);
            RefreshFiles();
        }

        void MoveSelected(int dir)
        {
            if (lvFiles.SelectedIndices.Count != 1) return;
            int i = lvFiles.SelectedIndices[0], j = i + dir;
            if (j < 0 || j >= tool.Files.Count) return;
            string tmp = tool.Files[i]; tool.Files[i] = tool.Files[j]; tool.Files[j] = tmp;
            RefreshFiles();
            lvFiles.SelectedIndices.Clear();
            lvFiles.Items[j].Selected = true;
            lvFiles.EnsureVisible(j);
        }

        void UpdateButtons()
        {
            bool idle = running == null;
            int nsel = lvFiles.SelectedIndices.Count;
            lstTools.Enabled = idle;
            btnAdd.Enabled = btnBrowse.Enabled = idle;
            btnRemove.Enabled = idle && nsel > 0;
            btnUp.Enabled = idle && nsel == 1 && lvFiles.SelectedIndices[0] > 0;
            btnDown.Enabled = idle && nsel == 1 && lvFiles.SelectedIndices[0] < tool.Files.Count - 1;
            btnClear.Enabled = idle && tool.Files.Count > 0;
            btnRun.Enabled = idle && tool.Files.Count >= tool.MinFiles;
            btnCancel.Enabled = !idle;
            btnOpen.Enabled = idle && lastResult != null;
            pnlOptions.Enabled = rbSame.Enabled = rbOther.Enabled = txtOut.Enabled = idle;
            progress.Visible = !idle;
        }

        void Log(string s)
        {
            if (InvokeRequired) { BeginInvoke(new Action<string>(Log), s); return; }
            txtLog.AppendText(s + "\r\n");
        }

        void BrowseOut()
        {
            using (var d = new FolderBrowserDialog { SelectedPath = txtOut.Text, Description = "Chọn thư mục lưu kết quả" })
                if (d.ShowDialog(this) == DialogResult.OK) { txtOut.Text = d.SelectedPath; rbOther.Checked = true; }
        }

        void OpenResult()
        {
            if (lastResult == null) return;
            try
            {
                if (File.Exists(lastResult)) Process.Start("explorer.exe", "/select," + Ext.Q(lastResult));
                else if (Directory.Exists(lastResult)) Process.Start("explorer.exe", Ext.Q(lastResult));
            }
            catch (Exception ex) { MessageBox.Show(this, ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }

        void CheckTools()
        {
            var missing = new List<string>();
            if (Ext.Qpdf == null) missing.Add("qpdf (gộp, tách, xoay, mật khẩu...)");
            if (Ext.Gs == null) missing.Add("Ghostscript (nén, PDF sang ảnh, OCR...)");
            if (Ext.Tesseract == null) missing.Add("Tesseract (OCR)");
            if (missing.Count == 0)
                Log("Sẵn sàng. Chọn công cụ bên trái, thêm file rồi bấm THỰC HIỆN.");
            else
                Log("Thiếu công cụ: " + string.Join(", ", missing) + ".\r\nHãy cài PDF24 Creator, hoặc chép các thư mục qpdf, gs, tesseract vào thư mục \"tools\" cạnh PdfTool.exe.");
        }

        // ------------------------------------------------------------------
        void RunTool()
        {
            if (tool.Files.Count < tool.MinFiles) return;
            string outDir = null;
            if (rbOther.Checked)
            {
                outDir = txtOut.Text.Trim();
                if (outDir.Length == 0 || !Path.IsPathRooted(outDir))
                {
                    MessageBox.Show(this, "Hãy chọn thư mục lưu kết quả.", "PdfTool", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            Action<Job> work = tool.Prepare();
            if (work == null) return;

            var job = new Job { Inputs = new List<string>(tool.Files), OutDir = outDir, Log = Log };
            running = job;
            lastResult = null;
            txtLog.Clear();
            Log("=== " + tool.Name + " ===");
            UpdateButtons();
            var started = DateTime.Now;

            var th = new Thread(() =>
            {
                try { work(job); }
                catch (Exception ex) { if (!job.Cancelled) { job.Errors++; Log("✖ " + ex.Message); } }
                finally { job.Cleanup(); }
                BeginInvoke(new Action(() =>
                {
                    running = null;
                    lastResult = job.Outputs.LastOrDefault();
                    string secs = (DateTime.Now - started).TotalSeconds.ToString("0.#");
                    if (job.Cancelled) Log("Đã dừng.");
                    else if (job.Errors == 0) Log("Hoàn tất (" + secs + " giây).");
                    else Log("Xong, có " + job.Errors + " lỗi (" + secs + " giây).");
                    UpdateButtons();
                    if (!job.Cancelled && job.Errors > 0)
                        MessageBox.Show(this, "Có " + job.Errors + " file bị lỗi. Xem chi tiết trong khung nhật ký.", "PdfTool", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }));
            });
            th.IsBackground = true;
            worker = th;
            th.Start();
        }

        // ------------------------------------------------------------------
        void LoadSettings()
        {
            txtOut.Text = Settings.Get("outdir", "");
            rbOther.Checked = Settings.Get("useoutdir", "0") == "1";
            var t = tools.FirstOrDefault(x => x.Name == Settings.Get("tool", ""));
            if (t != null) lstTools.SelectedItem = t;
        }

        public void SaveSettings()
        {
            Settings.Set("outdir", txtOut.Text);
            Settings.Set("useoutdir", rbOther.Checked ? "1" : "0");
            Settings.Set("tool", tool != null ? tool.Name : "");
        }

        // Nhận file từ thẻ "Soạn PDF"
        public void AddPdf(string path)
        {
            if (tool.ImageInput) lstTools.SelectedIndex = 0;
            AddFiles(new[] { path });
        }
    }

    // ------------------------------------------------------------------
    // Cài đặt dùng chung (%AppData%\PdfTool\settings.ini)
    // ------------------------------------------------------------------
    static class Settings
    {
        static readonly Dictionary<string, string> values = new Dictionary<string, string>();

        static string FilePath
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "PdfTool", "settings.ini"); }
        }

        public static void Load()
        {
            try
            {
                if (!File.Exists(FilePath)) return;
                foreach (string line in File.ReadAllLines(FilePath, Encoding.UTF8))
                {
                    int eq = line.IndexOf('=');
                    if (eq > 0) values[line.Substring(0, eq)] = line.Substring(eq + 1);
                }
            }
            catch { }
        }

        public static void Save()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
                File.WriteAllLines(FilePath, values.Select(kv => kv.Key + "=" + kv.Value), Encoding.UTF8);
            }
            catch { }
        }

        public static string Get(string key, string def) { string v; return values.TryGetValue(key, out v) ? v : def; }
        public static void Set(string key, string value) { values[key] = value ?? ""; }
    }

    // ------------------------------------------------------------------
    // Biểu tượng file/thư mục của Windows
    // ------------------------------------------------------------------
    static class ShellIcons
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        struct SHFILEINFO
        {
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)] public string szTypeName;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        static extern IntPtr SHGetFileInfo(string path, uint attr, ref SHFILEINFO info, uint size, uint flags);
        [DllImport("user32.dll")] static extern bool DestroyIcon(IntPtr h);

        const uint SHGFI_ICON = 0x100, SHGFI_SMALLICON = 0x1, SHGFI_USEFILEATTRIBUTES = 0x10;
        const uint ATTR_DIRECTORY = 0x10, ATTR_NORMAL = 0x80;

        // exact = true: lấy biểu tượng riêng của đường dẫn (ổ đĩa, Desktop...); ngược lại theo loại file
        public static string Key(ImageList list, string path, bool isDir, bool exact)
        {
            string key = exact ? path.ToLowerInvariant() : isDir ? "<dir>" : Path.GetExtension(path).ToLowerInvariant();
            if (list.Images.ContainsKey(key)) return key;
            var info = new SHFILEINFO();
            uint flags = SHGFI_ICON | SHGFI_SMALLICON | (exact ? 0 : SHGFI_USEFILEATTRIBUTES);
            string query = exact ? path : isDir ? "folder" : "file" + Path.GetExtension(path);
            SHGetFileInfo(query, isDir ? ATTR_DIRECTORY : ATTR_NORMAL, ref info, (uint)Marshal.SizeOf(info), flags);
            if (info.hIcon == IntPtr.Zero) return null;
            using (var ic = Icon.FromHandle(info.hIcon)) list.Images.Add(key, ic.ToBitmap());
            DestroyIcon(info.hIcon);
            return key;
        }
    }

    // ------------------------------------------------------------------
    // Hộp nhập liệu đơn giản
    // ------------------------------------------------------------------
    static class InputDialog
    {
        public static string Show(IWin32Window owner, string title, string prompt, string value, bool password)
        {
            using (var f = new Form())
            {
                float s;
                using (var g = f.CreateGraphics()) s = g.DpiX / 96f;
                f.Text = title;
                f.Font = new Font("Segoe UI", 9f);
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.StartPosition = FormStartPosition.CenterParent;
                f.MinimizeBox = f.MaximizeBox = false;
                f.ShowInTaskbar = false;
                f.AutoSize = true;
                f.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                f.Padding = new Padding((int)(12 * s));
                var flow = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false, Dock = DockStyle.Fill };
                flow.Controls.Add(new Label { Text = prompt, AutoSize = true, MaximumSize = new Size((int)(380 * s), 0) });
                var tb = new TextBox { Text = value, Width = (int)(380 * s), UseSystemPasswordChar = password, Margin = new Padding(0, (int)(8 * s), 0, (int)(8 * s)) };
                flow.Controls.Add(tb);
                var btns = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Width = (int)(380 * s), Margin = new Padding(0) };
                var cancel = new Button { Text = "Hủy", DialogResult = DialogResult.Cancel, Width = (int)(80 * s), Height = (int)(28 * s) };
                var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Width = (int)(80 * s), Height = (int)(28 * s) };
                btns.Controls.Add(cancel);
                btns.Controls.Add(ok);
                flow.Controls.Add(btns);
                f.Controls.Add(flow);
                f.AcceptButton = ok;
                f.CancelButton = cancel;
                return f.ShowDialog(owner) == DialogResult.OK ? tb.Text : null;
            }
        }
    }

    // ------------------------------------------------------------------
    // Trình duyệt thư mục bên trái (kéo file từ đây sang vùng soạn)
    // ------------------------------------------------------------------
    class FileBrowser : UserControl
    {
        readonly float scale;
        int S(int v) { return (int)(v * scale); }
        readonly TreeView tree;
        readonly ListView list;
        readonly TextBox address;
        readonly ImageList icons;
        readonly SplitContainer split;
        readonly System.Windows.Forms.Timer refreshTimer;
        readonly Stack<string> history = new Stack<string>();
        FileSystemWatcher watcher;
        string current;
        public Action<string[]> OpenFiles;

        public string CurrentFolder { get { return current; } }

        public static bool Supported(string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            return ext == ".pdf" || ImageUtil.Extensions.Contains(ext);
        }

        public FileBrowser(float scale)
        {
            this.scale = scale;
            Font = new Font("Segoe UI", 9f);
            icons = new ImageList { ColorDepth = ColorDepth.Depth32Bit, ImageSize = new Size(S(16), S(16)) };

            var bar = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 4, AutoSize = true, Padding = new Padding(S(2)) };
            for (int i = 0; i < 3; i++) bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            var tip = new ToolTip();
            Func<string, string, EventHandler, Button> btn = (text, hint, click) =>
            {
                var b = new Button { Text = text, Width = S(32), Height = S(26), Margin = new Padding(S(1), S(2), S(1), S(2)), FlatStyle = FlatStyle.Flat };
                b.FlatAppearance.BorderSize = 0;
                b.Click += click;
                tip.SetToolTip(b, hint);
                return b;
            };
            bar.Controls.Add(btn("←", "Quay lại", (s, e) => GoBack()), 0, 0);
            bar.Controls.Add(btn("↑", "Lên thư mục cha", (s, e) => GoUp()), 1, 0);
            bar.Controls.Add(btn("⟳", "Làm mới (F5)", (s, e) => Navigate(current, false)), 2, 0);
            address = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(S(3), S(4), S(3), S(3)) };
            address.KeyDown += (s, e) =>
            {
                if (e.KeyCode != Keys.Enter) return;
                e.SuppressKeyPress = true;
                string p = address.Text.Trim().Trim('"');
                if (Directory.Exists(p)) Navigate(p, true);
                else if (File.Exists(p) && Supported(p) && OpenFiles != null) OpenFiles(new[] { p });
            };
            bar.Controls.Add(address, 3, 0);

            split = new SplitContainer { Dock = DockStyle.Fill, FixedPanel = FixedPanel.Panel1 };
            tree = new TreeView { Dock = DockStyle.Fill, ImageList = icons, HideSelection = false, BorderStyle = BorderStyle.None, ShowLines = false, ItemHeight = S(22) };
            tree.BeforeExpand += (s, e) => Populate(e.Node);
            tree.AfterSelect += (s, e) => { if (e.Action != TreeViewAction.Unknown && e.Node.Tag is string) Navigate((string)e.Node.Tag, true); };

            list = new ListView
            {
                Dock = DockStyle.Fill, View = View.Details, SmallImageList = icons, FullRowSelect = true,
                HideSelection = false, BorderStyle = BorderStyle.None
            };
            list.Columns.Add("Tên", S(250));
            list.Columns.Add("Ngày sửa", S(115));
            list.Columns.Add("Kích thước", S(75), HorizontalAlignment.Right);
            list.ItemDrag += (s, e) =>
            {
                var paths = Selection();
                if (paths.Length > 0) list.DoDragDrop(new DataObject(DataFormats.FileDrop, paths), DragDropEffects.Copy);
            };
            list.DoubleClick += (s, e) => OpenSelected();
            list.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; OpenSelected(); }
                else if (e.KeyCode == Keys.Back) GoUp();
                else if (e.KeyCode == Keys.F5) Navigate(current, false);
                else if (e.Control && e.KeyCode == Keys.A) foreach (ListViewItem it in list.Items) it.Selected = true;
            };

            split.Panel1.Controls.Add(tree);
            split.Panel2.Controls.Add(list);
            Controls.Add(split);
            Controls.Add(bar);

            refreshTimer = new System.Windows.Forms.Timer { Interval = 600 };
            refreshTimer.Tick += (s, e) => { refreshTimer.Stop(); Navigate(current, false); };
            // Đặt độ rộng cây thư mục khi khung đã có kích thước thật
            bool sized = false;
            split.SizeChanged += (s, e) =>
            {
                if (sized || split.Width < S(400)) return;
                sized = true;
                try { split.SplitterDistance = S(190); } catch { }
            };
            AddRoots();
        }

        string[] Selection()
        {
            return list.SelectedItems.Cast<ListViewItem>().Select(i => i.Tag as string).Where(p => p != null).ToArray();
        }

        void OpenSelected()
        {
            var paths = Selection();
            if (paths.Length == 1 && Directory.Exists(paths[0])) { Navigate(paths[0], true); return; }
            var files = paths.Where(File.Exists).ToArray();
            if (files.Length > 0 && OpenFiles != null) OpenFiles(files);
        }

        void AddRoots()
        {
            string profile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            AddRoot("Desktop", Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));
            AddRoot("Tài liệu (Documents)", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            AddRoot("Tải về (Downloads)", Path.Combine(profile, "Downloads"));
            AddRoot("Ảnh (Pictures)", Environment.GetFolderPath(Environment.SpecialFolder.MyPictures));
            foreach (var d in DriveInfo.GetDrives())
            {
                try
                {
                    if (!d.IsReady) continue;
                    string label = d.VolumeLabel;
                    AddRoot((label.Length > 0 ? label + " " : "") + "(" + d.Name.TrimEnd('\\') + ")", d.RootDirectory.FullName);
                }
                catch { }
            }
        }

        void AddRoot(string text, string path)
        {
            if (string.IsNullOrEmpty(path) || !Directory.Exists(path)) return;
            string key = ShellIcons.Key(icons, path, true, true);
            var n = new TreeNode(text) { Tag = path, ImageKey = key, SelectedImageKey = key };
            n.Nodes.Add(new TreeNode("..."));
            tree.Nodes.Add(n);
        }

        void Populate(TreeNode node)
        {
            if (node.Nodes.Count != 1 || node.Nodes[0].Tag != null) return;
            node.Nodes.Clear();
            string key = ShellIcons.Key(icons, "", true, false);
            foreach (var d in SubDirs((string)node.Tag))
            {
                var n = new TreeNode(d.Name) { Tag = d.FullName, ImageKey = key, SelectedImageKey = key };
                n.Nodes.Add(new TreeNode("..."));
                node.Nodes.Add(n);
            }
        }

        static bool IsShown(FileSystemInfo f) { return (f.Attributes & (FileAttributes.Hidden | FileAttributes.System)) == 0; }

        static List<DirectoryInfo> SubDirs(string path)
        {
            try
            {
                return new DirectoryInfo(path).GetDirectories().Where(IsShown)
                    .OrderBy(d => d.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
            }
            catch { return new List<DirectoryInfo>(); }
        }

        void GoUp()
        {
            if (current == null) return;
            var parent = Directory.GetParent(current);
            if (parent != null) Navigate(parent.FullName, true);
        }

        void GoBack()
        {
            if (history.Count > 0) Navigate(history.Pop(), false);
        }

        public void Navigate(string path, bool addHistory)
        {
            if (string.IsNullOrEmpty(path) || !Directory.Exists(path)) return;
            bool refresh = string.Equals(current, path, StringComparison.OrdinalIgnoreCase);
            if (addHistory && current != null && !refresh) history.Push(current);
            var keep = refresh ? new HashSet<string>(Selection(), StringComparer.OrdinalIgnoreCase) : new HashSet<string>();
            int top = refresh && list.TopItem != null ? list.TopItem.Index : 0;
            current = path;
            address.Text = path;
            list.BeginUpdate();
            list.Items.Clear();
            try
            {
                var di = new DirectoryInfo(path);
                string dirKey = ShellIcons.Key(icons, "", true, false);
                foreach (var d in SubDirs(path))
                    list.Items.Add(new ListViewItem(new[] { d.Name, d.LastWriteTime.ToString("dd/MM/yyyy HH:mm"), "" })
                        { Tag = d.FullName, ImageKey = dirKey, Selected = keep.Contains(d.FullName) });
                foreach (var f in di.GetFiles().Where(f => IsShown(f) && Supported(f.Name)).OrderBy(f => f.Name, StringComparer.CurrentCultureIgnoreCase))
                    list.Items.Add(new ListViewItem(new[] { f.Name, f.LastWriteTime.ToString("dd/MM/yyyy HH:mm"), Job.SizeText(f.Length) })
                        { Tag = f.FullName, ImageKey = ShellIcons.Key(icons, f.FullName, false, false), Selected = keep.Contains(f.FullName) });
            }
            catch (Exception ex) { list.Items.Add(new ListViewItem("(" + ex.Message + ")")); }
            list.EndUpdate();
            if (refresh && top > 0 && top < list.Items.Count) list.TopItem = list.Items[top];
            Watch(path);
            Settings.Set("browser", path);
        }

        // Tự làm mới khi thư mục thay đổi (vd. vừa lưu file vào đây)
        void Watch(string path)
        {
            try
            {
                if (watcher == null)
                {
                    watcher = new FileSystemWatcher
                    {
                        SynchronizingObject = this,
                        NotifyFilter = NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.LastWrite | NotifyFilters.Size
                    };
                    FileSystemEventHandler h = (s, e) => { refreshTimer.Stop(); refreshTimer.Start(); };
                    watcher.Created += h;
                    watcher.Deleted += h;
                    watcher.Changed += h;
                    watcher.Renamed += (s, e) => h(s, e);
                }
                watcher.EnableRaisingEvents = false;
                watcher.Path = path;
                watcher.EnableRaisingEvents = true;
            }
            catch { }
        }
    }

    // ------------------------------------------------------------------
    // Mô hình tài liệu đang soạn
    // ------------------------------------------------------------------
    class PageRef
    {
        public string File;   // bản sao tạm của file nguồn
        public int Index;     // số trang trong file nguồn (từ 1)
        public int Rot;       // góc xoay thêm: 0, 90, 180, 270
        public string Key { get { return File + "|" + Index; } }
        public PageRef Clone() { return new PageRef { File = File, Index = Index, Rot = Rot }; }
    }

    class PdfDoc
    {
        public string Name;
        public string FilePath;   // file đã mở/lưu gần nhất (null nếu là tài liệu mới)
        public bool Modified;
        public readonly List<PageRef> Pages = new List<PageRef>();
    }

    // Phiên làm việc: giữ bản sao tạm của các file đã mở, tạo ảnh thu nhỏ, lưu tài liệu
    class Session : IDisposable
    {
        public readonly string Dir = Path.Combine(Path.GetTempPath(), "pdftool_" + Guid.NewGuid().ToString("N"));
        readonly Dictionary<string, Bitmap> thumbs = new Dictionary<string, Bitmap>();
        readonly System.Collections.Concurrent.BlockingCollection<KeyValuePair<string, int>> queue =
            new System.Collections.Concurrent.BlockingCollection<KeyValuePair<string, int>>();
        public Action ThumbsChanged;     // gọi trên luồng nền
        const int ThumbDpi = 30;

        public Session()
        {
            Directory.CreateDirectory(Dir);
            var t = new Thread(ThumbWorker) { IsBackground = true, Priority = ThreadPriority.BelowNormal };
            t.Start();
        }

        public Bitmap Thumb(PageRef p)
        {
            Bitmap b;
            lock (thumbs) return thumbs.TryGetValue(p.Key, out b) ? b : null;
        }

        void ThumbWorker()
        {
            foreach (var item in queue.GetConsumingEnumerable())
            {
                string gs = Ext.Gs;
                if (gs == null) continue;
                string file = item.Key;
                string tmp = Path.Combine(Dir, "thumbs_" + Guid.NewGuid().ToString("N"));
                try { Directory.CreateDirectory(tmp); } catch { continue; }
                for (int first = 1; first <= item.Value; first += 6)
                {
                    if (!File.Exists(file)) break;
                    int last = Math.Min(item.Value, first + 5);
                    try
                    {
                        string output;
                        Ext.Run(gs, "-q -dSAFER -dBATCH -dNOPAUSE -sDEVICE=png16m -r" + ThumbDpi +
                                    " -dTextAlphaBits=4 -dGraphicsAlphaBits=4 -dFirstPage=" + first + " -dLastPage=" + last + " " +
                                    Ext.Q("-sOutputFile=" + Path.Combine(tmp, "t-%d.png")) + " " + Ext.Q(file), out output);
                        for (int i = first; i <= last; i++)
                        {
                            string png = Path.Combine(tmp, "t-" + (i - first + 1) + ".png");
                            if (!File.Exists(png)) continue;
                            Bitmap bmp;
                            using (var b = new Bitmap(png)) bmp = new Bitmap(b);
                            File.Delete(png);
                            lock (thumbs) thumbs[file + "|" + i] = bmp;
                        }
                    }
                    catch { }
                    var cb = ThumbsChanged;
                    if (cb != null) cb();
                }
                try { Directory.Delete(tmp, true); } catch { }
            }
        }

        // Đọc file PDF/ảnh vào phiên; trả về null nếu người dùng hủy nhập mật khẩu
        public PdfDoc Import(string path, IWin32Window owner)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            string copy = Path.Combine(Dir, Guid.NewGuid().ToString("N") + ".pdf");
            bool isPdf = ext == ".pdf";
            if (!isPdf)
            {
                var pages = new List<PdfPageSpec>();
                foreach (PdfImage im in ImageUtil.Load(path))
                {
                    double w = im.W * 72.0 / im.Dpi, h = im.H * 72.0 / im.Dpi;
                    var box = new PageBox { X1 = w, Y1 = h };
                    var spec = new PdfPageSpec { Box = box };
                    spec.Images.Add(new KeyValuePair<PdfImage, double[]>(im, box.Place(0, 0, w, h)));
                    pages.Add(spec);
                }
                PdfWriter.Save(pages, copy);
            }
            else
            {
                string qpdf = Ext.Require(Ext.Qpdf, "qpdf");
                string output;
                int enc = Ext.Run(qpdf, "--requires-password " + Ext.Q(path), out output);
                if (enc == 0) // cần mật khẩu để mở
                {
                    while (true)
                    {
                        string pw = InputDialog.Show(owner, "Mật khẩu", "File \"" + Path.GetFileName(path) + "\" được bảo vệ. Nhập mật khẩu để mở:", "", true);
                        if (pw == null) return null;
                        if (Ext.Run(qpdf, "--warning-exit-0 " + Ext.Q("--password=" + pw) + " --decrypt " + Ext.Q(path) + " " + Ext.Q(copy), out output) == 0) break;
                        MessageBox.Show(owner, "Sai mật khẩu.", "PdfTool", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else if (enc == 3) // chỉ bị hạn chế quyền
                    Ext.QpdfRun("--decrypt " + Ext.Q(path) + " " + Ext.Q(copy));
                else
                    File.Copy(path, copy);
            }
            string npages = Ext.QpdfRun("--show-npages " + Ext.Q(copy)).Trim();
            int n = int.Parse(npages.Split('\n').Last().Trim());
            var doc = new PdfDoc { Name = Path.GetFileNameWithoutExtension(path), FilePath = isPdf ? path : null, Modified = !isPdf };
            for (int i = 1; i <= n; i++) doc.Pages.Add(new PageRef { File = copy, Index = i });
            queue.Add(new KeyValuePair<string, int>(copy, n));
            return doc;
        }

        static string Ranges(List<int> nums)
        {
            var parts = new List<string>();
            for (int i = 0; i < nums.Count; )
            {
                int j = i;
                while (j + 1 < nums.Count && nums[j + 1] == nums[j] + 1) j++;
                parts.Add(j > i ? nums[i] + "-" + nums[j] : nums[i].ToString());
                i = j + 1;
            }
            return string.Join(",", parts);
        }

        public void Save(PdfDoc doc, string outPath)
        {
            var args = new List<string> { "--warning-exit-0", "--empty", "--pages" };
            for (int i = 0; i < doc.Pages.Count; )
            {
                string f = doc.Pages[i].File;
                var nums = new List<int>();
                while (i < doc.Pages.Count && doc.Pages[i].File == f) nums.Add(doc.Pages[i++].Index);
                args.Add(f);
                args.Add(Ranges(nums));
            }
            args.Add("--");
            foreach (int r in new[] { 90, 180, 270 })
            {
                var idx = Enumerable.Range(0, doc.Pages.Count).Where(i => doc.Pages[i].Rot == r).Select(i => i + 1).ToList();
                if (idx.Count > 0) args.Add("--rotate=+" + r + ":" + Ranges(idx));
            }
            string id = Guid.NewGuid().ToString("N");
            string argFile = Path.Combine(Dir, "args_" + id + ".txt");
            string tmpOut = Path.Combine(Dir, "save_" + id + ".pdf");
            File.WriteAllLines(argFile, args, new UTF8Encoding(false));
            try
            {
                // Ghi ra file tạm trước để không hỏng file đích nếu có lỗi
                Ext.Check(Ext.Require(Ext.Qpdf, "qpdf"), Ext.Q("@" + argFile) + " " + Ext.Q(tmpOut));
                File.Copy(tmpOut, outPath, true);
            }
            finally
            {
                try { File.Delete(argFile); } catch { }
                try { File.Delete(tmpOut); } catch { }
            }
        }

        public Bitmap Render(PageRef p, int dpi)
        {
            string png = Path.Combine(Dir, "prev_" + Guid.NewGuid().ToString("N") + ".png");
            Ext.Check(Ext.Require(Ext.Gs, "Ghostscript"), "-q -dSAFER -dBATCH -dNOPAUSE -sDEVICE=png16m -r" + dpi +
                      " -dTextAlphaBits=4 -dGraphicsAlphaBits=4 -dFirstPage=" + p.Index + " -dLastPage=" + p.Index + " " +
                      Ext.Q("-sOutputFile=" + png) + " " + Ext.Q(p.File));
            Bitmap bmp;
            using (var b = new Bitmap(png)) bmp = new Bitmap(b);
            try { File.Delete(png); } catch { }
            if (p.Rot == 90) bmp.RotateFlip(RotateFlipType.Rotate90FlipNone);
            else if (p.Rot == 180) bmp.RotateFlip(RotateFlipType.Rotate180FlipNone);
            else if (p.Rot == 270) bmp.RotateFlip(RotateFlipType.Rotate270FlipNone);
            return bmp;
        }

        public void Dispose()
        {
            queue.CompleteAdding();
            try { Directory.Delete(Dir, true); } catch { }
        }
    }

    // ------------------------------------------------------------------
    // Vùng soạn: hiển thị tài liệu dạng ảnh thu nhỏ, kéo thả để sắp xếp
    // ------------------------------------------------------------------
    class Workspace : Panel
    {
        public readonly List<PdfDoc> Docs = new List<PdfDoc>();
        public readonly HashSet<PageRef> Selected = new HashSet<PageRef>();
        public PdfDoc Active;
        public Session Session;
        public event Action Changed;
        public Action<PdfDoc> SaveRequested, CloseRequested;
        public Action<string[], PdfDoc, int> FilesDropped;     // doc null = tạo tài liệu mới
        public Action<PageRef> PreviewRequested;
        public ContextMenuStrip PageMenu, DocMenu;
        int newCounter;

        readonly float scale;
        int box;
        int S(int v) { return (int)(v * scale); }

        class DocLayout
        {
            public PdfDoc Doc;
            public Rectangle Block, Header, SaveBtn, CloseBtn;
            public readonly List<Rectangle> Cells = new List<Rectangle>();
        }
        readonly List<DocLayout> layout = new List<DocLayout>();

        Point downPt;
        bool dragArmed;
        PageRef pendingSingle, anchor;
        bool dropActive;
        PdfDoc dropDoc;
        int dropIndex;
        const string PagesFormat = "PdfToolPages";

        static readonly Color Bg = Color.FromArgb(43, 43, 43), BlockBg = Color.FromArgb(54, 54, 56), HeaderBg = Color.FromArgb(64, 64, 68),
                              Accent = Color.FromArgb(0, 120, 212), TextC = Color.FromArgb(235, 235, 235), SubText = Color.FromArgb(165, 165, 165);

        public Workspace(float scale)
        {
            this.scale = scale;
            box = S(150);
            SetStyle(ControlStyles.Selectable | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
            AutoScroll = true;
            AllowDrop = true;
            TabStop = true;
            BackColor = Bg;
        }

        void OnChanged()
        {
            var h = Changed;
            if (h != null) h();
        }

        int CellW { get { return box + S(26); } }
        int CellH { get { return box + S(34); } }

        public void Relayout()
        {
            if (Active != null && !Docs.Contains(Active)) Active = Docs.FirstOrDefault();
            Selected.RemoveWhere(p => !Docs.Any(d => d.Pages.Contains(p)));
            layout.Clear();
            int m = S(14), y = m;
            int width = Math.Max(CellW + 2 * m + S(16), ClientSize.Width);
            int cols = Math.Max(1, (width - 2 * m - S(16)) / CellW);
            int hh = S(40);
            foreach (var d in Docs)
            {
                var L = new DocLayout { Doc = d };
                int rows = Math.Max(1, (d.Pages.Count + cols - 1) / cols);
                L.Block = new Rectangle(m, y, width - 2 * m, hh + S(8) + rows * CellH + S(6));
                L.Header = new Rectangle(m, y, width - 2 * m, hh);
                int bh = S(26);
                L.CloseBtn = new Rectangle(L.Header.Right - S(10) - S(30), y + (hh - bh) / 2, S(30), bh);
                L.SaveBtn = new Rectangle(L.CloseBtn.Left - S(6) - S(64), y + (hh - bh) / 2, S(64), bh);
                for (int i = 0; i < d.Pages.Count; i++)
                    L.Cells.Add(new Rectangle(m + S(8) + (i % cols) * CellW, y + hh + S(8) + (i / cols) * CellH, CellW, CellH));
                layout.Add(L);
                y = L.Block.Bottom + m;
            }
            AutoScrollMinSize = new Size(0, y);
            Invalidate();
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            Relayout();
        }

        // ---------------- vẽ ----------------
        static GraphicsPath Round(Rectangle r, int rad)
        {
            var p = new GraphicsPath();
            int d = Math.Max(1, rad * 2);
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Bg);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;

            if (Docs.Count == 0) { DrawEmpty(g); if (dropActive) DrawNewDocTarget(g); return; }

            var state = g.Save();
            g.TranslateTransform(AutoScrollPosition.X, AutoScrollPosition.Y);
            var clip = e.ClipRectangle;
            clip.Offset(-AutoScrollPosition.X, -AutoScrollPosition.Y);

            using (var titleFont = new Font("Segoe UI Semibold", 10.5f))
            using (var small = new Font("Segoe UI", 8.5f))
            using (var sfC = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
            using (var sfL = new StringFormat { LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter, FormatFlags = StringFormatFlags.NoWrap })
            {
                foreach (var L in layout)
                {
                    if (!L.Block.IntersectsWith(clip)) continue;
                    var d = L.Doc;
                    bool active = d == Active;
                    using (var path = Round(L.Block, S(6)))
                    using (var b = new SolidBrush(BlockBg)) g.FillPath(b, path);
                    using (var path = Round(L.Header, S(6)))
                    using (var b = new SolidBrush(active ? Color.FromArgb(72, 72, 78) : HeaderBg)) g.FillPath(b, path);
                    if (active) using (var b = new SolidBrush(Accent)) g.FillRectangle(b, L.Header.X, L.Header.Y + S(8), S(4), L.Header.Height - S(16));

                    var titleRect = new RectangleF(L.Header.X + S(14), L.Header.Y, L.SaveBtn.Left - L.Header.X - S(24), L.Header.Height);
                    string title = d.Name + (d.Modified ? " *" : "");
                    string sub = "   " + d.Pages.Count + " trang";
                    using (var tb = new SolidBrush(TextC)) g.DrawString(title, titleFont, tb, titleRect, sfL);
                    var tsz = g.MeasureString(title, titleFont);
                    if (tsz.Width + S(80) < titleRect.Width)
                        using (var sb = new SolidBrush(SubText))
                            g.DrawString(sub, small, sb, new RectangleF(titleRect.X + tsz.Width, titleRect.Y, S(120), titleRect.Height), sfL);

                    DrawButton(g, L.SaveBtn, "Lưu", small, sfC, d.Modified);
                    DrawButton(g, L.CloseBtn, "✕", small, sfC, false);

                    for (int i = 0; i < d.Pages.Count; i++)
                        if (L.Cells[i].IntersectsWith(clip)) DrawPage(g, d.Pages[i], L.Cells[i], i + 1, small, sfC);
                }

                // Vị trí thả
                if (dropActive && dropDoc != null)
                {
                    var L = layout.FirstOrDefault(x => x.Doc == dropDoc);
                    if (L != null && L.Cells.Count > 0)
                    {
                        Rectangle c = dropIndex < L.Cells.Count ? L.Cells[dropIndex] : L.Cells[L.Cells.Count - 1];
                        int x = dropIndex < L.Cells.Count ? c.Left : c.Right;
                        using (var pen = new Pen(Accent, S(4))) g.DrawLine(pen, x, c.Top + S(2), x, c.Top + S(8) + box);
                    }
                }
            }
            g.Restore(state);
            if (dropActive && dropDoc == null) DrawNewDocTarget(g);
        }

        void DrawButton(Graphics g, Rectangle r, string text, Font f, StringFormat sf, bool highlight)
        {
            using (var path = Round(r, S(4)))
            {
                using (var b = new SolidBrush(highlight ? Accent : Color.FromArgb(90, 90, 96))) g.FillPath(b, path);
                g.DrawString(text, f, Brushes.White, r, sf);
            }
        }

        void DrawPage(Graphics g, PageRef p, Rectangle cell, int number, Font f, StringFormat sf)
        {
            bool sel = Selected.Contains(p);
            if (sel)
            {
                var r0 = Rectangle.Inflate(cell, -S(3), -S(2));
                using (var path = Round(r0, S(5)))
                {
                    using (var b = new SolidBrush(Color.FromArgb(80, 0, 120, 212))) g.FillPath(b, path);
                    using (var pen = new Pen(Accent, S(2))) g.DrawPath(pen, path);
                }
            }
            Bitmap bmp = Session.Thumb(p);
            double w = bmp != null ? bmp.Width : 210, h = bmp != null ? bmp.Height : 297;
            bool swap = p.Rot % 180 != 0;
            double vw = swap ? h : w, vh = swap ? w : h;
            double s = Math.Min(box / vw, box / vh);
            int dw = Math.Max(4, (int)(vw * s)), dh = Math.Max(4, (int)(vh * s));
            var r = new Rectangle(cell.X + (cell.Width - dw) / 2, cell.Y + S(8) + (box - dh) / 2, dw, dh);
            using (var shadow = new SolidBrush(Color.FromArgb(90, 0, 0, 0))) g.FillRectangle(shadow, r.X + S(2), r.Y + S(3), r.Width, r.Height);
            g.FillRectangle(Brushes.White, r);
            if (bmp != null)
            {
                var st = g.Save();
                g.TranslateTransform(r.X + r.Width / 2f, r.Y + r.Height / 2f);
                g.RotateTransform(p.Rot);
                float uw = swap ? r.Height : r.Width, uh = swap ? r.Width : r.Height;
                g.DrawImage(bmp, -uw / 2, -uh / 2, uw, uh);
                g.Restore(st);
            }
            else g.DrawString("...", f, Brushes.Gray, r, sf);
            using (var b = new SolidBrush(sel ? Color.White : SubText))
                g.DrawString(number.ToString(), f, b, new Rectangle(cell.X, cell.Y + S(8) + box + S(2), cell.Width, S(20)), sf);
        }

        void DrawEmpty(Graphics g)
        {
            var r = Rectangle.Inflate(ClientRectangle, -S(40), -S(40));
            if (r.Width < 50 || r.Height < 50) return;
            using (var pen = new Pen(Color.FromArgb(90, 90, 90), S(2)) { DashStyle = DashStyle.Dash })
            using (var path = Round(r, S(12))) g.DrawPath(pen, path);
            using (var big = new Font("Segoe UI Semibold", 15f))
            using (var small = new Font("Segoe UI", 10f))
            using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
            using (var b1 = new SolidBrush(TextC))
            using (var b2 = new SolidBrush(SubText))
            {
                int cy = r.Y + r.Height / 2;
                g.DrawString("Kéo thả tài liệu PDF hoặc ảnh vào đây để chỉnh sửa", big, b1, new Rectangle(r.X, cy - S(40), r.Width, S(40)), sf);
                g.DrawString("hoặc nhấn đúp vào file ở danh sách bên trái, hoặc bấm \"Mở...\"", small, b2, new Rectangle(r.X, cy + S(4), r.Width, S(30)), sf);
            }
        }

        void DrawNewDocTarget(Graphics g)
        {
            var r = Rectangle.Inflate(ClientRectangle, -S(4), -S(4));
            using (var pen = new Pen(Accent, S(3)) { DashStyle = DashStyle.Dash }) g.DrawRectangle(pen, r);
            var lbl = new Rectangle(r.X + r.Width / 2 - S(130), r.Bottom - S(44), S(260), S(32));
            using (var path = Round(lbl, S(6)))
            using (var b = new SolidBrush(Accent)) g.FillPath(b, path);
            using (var f = new Font("Segoe UI", 9.5f))
            using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                g.DrawString("Thả để tạo tài liệu mới", f, Brushes.White, lbl, sf);
        }

        // ---------------- chuột ----------------
        Point ToContent(Point client) { return new Point(client.X - AutoScrollPosition.X, client.Y - AutoScrollPosition.Y); }
        DocLayout DocAt(Point p) { return layout.FirstOrDefault(L => L.Block.Contains(p)); }

        static int PageAt(DocLayout L, Point p)
        {
            for (int i = 0; i < L.Cells.Count; i++) if (L.Cells[i].Contains(p)) return i;
            return -1;
        }

        static int InsertIndex(DocLayout L, Point p)
        {
            for (int i = 0; i < L.Cells.Count; i++)
            {
                var c = L.Cells[i];
                if (p.Y < c.Top || (p.Y < c.Bottom && p.X < c.Left + c.Width / 2)) return i;
            }
            return L.Cells.Count;
        }

        public List<PageRef> OrderedSelection()
        {
            return Docs.SelectMany(d => d.Pages).Where(Selected.Contains).ToList();
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            Focus();
            Point p = ToContent(e.Location);
            var L = DocAt(p);
            pendingSingle = null;
            dragArmed = false;
            bool ctrl = (ModifierKeys & Keys.Control) != 0, shift = (ModifierKeys & Keys.Shift) != 0;
            if (L == null)
            {
                if (!ctrl) Selected.Clear();
                Invalidate();
                OnChanged();
                return;
            }
            Active = L.Doc;
            if (e.Button == MouseButtons.Left && L.SaveBtn.Contains(p)) { Invalidate(); OnChanged(); SaveRequested(L.Doc); return; }
            if (e.Button == MouseButtons.Left && L.CloseBtn.Contains(p)) { CloseRequested(L.Doc); return; }

            int i = PageAt(L, p);
            if (i < 0)
            {
                if (!ctrl) Selected.Clear();
                if (e.Clicks == 2 && L.Header.Contains(p)) foreach (var pg0 in L.Doc.Pages) Selected.Add(pg0);
                Invalidate();
                OnChanged();
                if (e.Button == MouseButtons.Right && DocMenu != null) DocMenu.Show(this, e.Location);
                return;
            }

            PageRef pg = L.Doc.Pages[i];
            if (e.Button == MouseButtons.Right)
            {
                if (!Selected.Contains(pg)) { Selected.Clear(); Selected.Add(pg); anchor = pg; }
            }
            else if (shift && anchor != null && L.Doc.Pages.Contains(anchor))
            {
                int a = L.Doc.Pages.IndexOf(anchor);
                if (!ctrl) Selected.Clear();
                for (int k = Math.Min(a, i); k <= Math.Max(a, i); k++) Selected.Add(L.Doc.Pages[k]);
            }
            else if (ctrl)
            {
                if (!Selected.Remove(pg)) Selected.Add(pg);
                anchor = pg;
            }
            else
            {
                if (Selected.Contains(pg)) pendingSingle = pg;
                else { Selected.Clear(); Selected.Add(pg); }
                anchor = pg;
            }
            Invalidate();
            OnChanged();
            if (e.Button == MouseButtons.Left)
            {
                if (e.Clicks == 2) { pendingSingle = null; PreviewRequested(pg); return; }
                dragArmed = Selected.Contains(pg);
                downPt = e.Location;
            }
            else if (e.Button == MouseButtons.Right && PageMenu != null) PageMenu.Show(this, e.Location);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (!dragArmed || e.Button != MouseButtons.Left) return;
            var ds = SystemInformation.DragSize;
            if (Math.Abs(e.X - downPt.X) < ds.Width && Math.Abs(e.Y - downPt.Y) < ds.Height) return;
            dragArmed = false;
            pendingSingle = null;
            DoDragDrop(new DataObject(PagesFormat, "pages"), DragDropEffects.Move | DragDropEffects.Copy);
            dropActive = false;
            Invalidate();
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            dragArmed = false;
            if (pendingSingle != null)
            {
                Selected.Clear();
                Selected.Add(pendingSingle);
                pendingSingle = null;
                Invalidate();
                OnChanged();
            }
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            if ((ModifierKeys & Keys.Control) != 0) { Zoom(e.Delta > 0 ? S(25) : -S(25)); return; }
            base.OnMouseWheel(e);
        }

        protected override bool IsInputKey(Keys keyData)
        {
            return keyData == Keys.Delete || base.IsInputKey(keyData);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.KeyCode == Keys.Delete) DeleteSelected();
            else if (e.Control && e.KeyCode == Keys.A) SelectAll();
            else if (e.KeyCode == Keys.Enter && Selected.Count > 0) PreviewRequested(OrderedSelection()[0]);
            else if (e.Control && (e.KeyCode == Keys.Oemplus || e.KeyCode == Keys.Add)) Zoom(S(25));
            else if (e.Control && (e.KeyCode == Keys.OemMinus || e.KeyCode == Keys.Subtract)) Zoom(-S(25));
        }

        // ---------------- kéo thả ----------------
        void UpdateDrop(DragEventArgs e)
        {
            bool pages = e.Data.GetDataPresent(PagesFormat), files = e.Data.GetDataPresent(DataFormats.FileDrop);
            if (!pages && !files) { e.Effect = DragDropEffects.None; return; }
            e.Effect = pages ? ((e.KeyState & 8) != 0 ? DragDropEffects.Copy : DragDropEffects.Move) : DragDropEffects.Copy;
            Point client = PointToClient(new Point(e.X, e.Y));
            int edge = S(40);
            if (client.Y < edge) ScrollBy(-S(25));
            else if (client.Y > ClientSize.Height - edge) ScrollBy(S(25));
            Point p = ToContent(client);
            var L = DocAt(p);
            PdfDoc d = L != null ? L.Doc : null;
            int idx = L != null ? InsertIndex(L, p) : 0;
            if (!dropActive || d != dropDoc || idx != dropIndex)
            {
                dropActive = true;
                dropDoc = d;
                dropIndex = idx;
                Invalidate();
            }
        }

        void ScrollBy(int dy)
        {
            AutoScrollPosition = new Point(-AutoScrollPosition.X, -AutoScrollPosition.Y + dy);
        }

        protected override void OnDragEnter(DragEventArgs e) { base.OnDragEnter(e); UpdateDrop(e); }
        protected override void OnDragOver(DragEventArgs e) { base.OnDragOver(e); UpdateDrop(e); }
        protected override void OnDragLeave(EventArgs e) { base.OnDragLeave(e); dropActive = false; Invalidate(); }

        protected override void OnDragDrop(DragEventArgs e)
        {
            base.OnDragDrop(e);
            dropActive = false;
            Invalidate();
            PdfDoc doc = dropDoc;
            int idx = dropIndex;
            if (e.Data.GetDataPresent(PagesFormat))
                MovePages(doc, idx, (e.KeyState & 8) != 0);
            else if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                // Xử lý sau khi kết thúc thao tác kéo để không làm treo cửa sổ nguồn (Explorer)
                BeginInvoke(new Action(() => FilesDropped(files, doc, idx)));
            }
        }

        // ---------------- thao tác ----------------
        PdfDoc NewDoc()
        {
            return new PdfDoc { Name = "Tài liệu mới " + (++newCounter), Modified = true };
        }

        public void MovePages(PdfDoc target, int index, bool copy)
        {
            var moving = OrderedSelection();
            if (moving.Count == 0) return;
            if (target != null && !Docs.Contains(target)) target = null;
            if (!copy && target != null && moving.All(target.Pages.Contains))
            {
                int first = target.Pages.IndexOf(moving[0]);
                bool contiguous = moving.Select((pg, k) => target.Pages.IndexOf(pg) == first + k).All(b => b);
                if (contiguous && index >= first && index <= first + moving.Count) return; // thả về chỗ cũ
            }
            if (copy) moving = moving.Select(pg => pg.Clone()).ToList();
            else
                foreach (var d in Docs)
                {
                    for (int i = d.Pages.Count - 1; i >= 0; i--)
                    {
                        if (!Selected.Contains(d.Pages[i])) continue;
                        if (d == target && i < index) index--;
                        d.Pages.RemoveAt(i);
                        d.Modified = true;
                    }
                }
            if (target == null)
            {
                target = NewDoc();
                Docs.Add(target);
                index = 0;
            }
            index = Math.Max(0, Math.Min(index, target.Pages.Count));
            target.Pages.InsertRange(index, moving);
            target.Modified = true;
            Docs.RemoveAll(d => d.Pages.Count == 0);
            Selected.Clear();
            foreach (var pg in moving) Selected.Add(pg);
            Active = target;
            Relayout();
            OnChanged();
        }

        public void AddDoc(PdfDoc doc, PdfDoc target, ref int index)
        {
            if (target != null && Docs.Contains(target))
            {
                index = Math.Max(0, Math.Min(index, target.Pages.Count));
                target.Pages.InsertRange(index, doc.Pages);
                index += doc.Pages.Count;
                target.Modified = true;
                Active = target;
            }
            else
            {
                Docs.Add(doc);
                Active = doc;
            }
        }

        public void RotateSelected(int deg)
        {
            var sel = OrderedSelection();
            foreach (var pg in sel) pg.Rot = ((pg.Rot + deg) % 360 + 360) % 360;
            foreach (var d in Docs) if (d.Pages.Any(Selected.Contains)) d.Modified = true;
            Relayout();
            OnChanged();
        }

        public void DeleteSelected()
        {
            if (Selected.Count == 0) return;
            foreach (var d in Docs)
                if (d.Pages.RemoveAll(Selected.Contains) > 0) d.Modified = true;
            Docs.RemoveAll(d => d.Pages.Count == 0);
            Selected.Clear();
            Relayout();
            OnChanged();
        }

        public void ExtractSelected()
        {
            if (Selected.Count > 0) MovePages(null, 0, false);
        }

        public void SelectAll()
        {
            Selected.Clear();
            foreach (var pg in Active != null ? Active.Pages : Docs.SelectMany(d => d.Pages).ToList()) Selected.Add(pg);
            Invalidate();
            OnChanged();
        }

        public void MergeAll()
        {
            if (Docs.Count < 2) return;
            var first = Docs[0];
            foreach (var d in Docs.Skip(1)) first.Pages.AddRange(d.Pages);
            Docs.RemoveRange(1, Docs.Count - 1);
            first.Modified = true;
            Active = first;
            Relayout();
            OnChanged();
        }

        public void RemoveDoc(PdfDoc d)
        {
            Docs.Remove(d);
            Relayout();
            OnChanged();
        }

        public void Zoom(int delta)
        {
            box = Math.Max(S(80), Math.Min(S(360), box + delta));
            Relayout();
        }
    }

    // ------------------------------------------------------------------
    // Xem trước trang lớn
    // ------------------------------------------------------------------
    class PreviewForm : Form
    {
        readonly Session session;
        readonly List<PageRef> pages;
        int index;
        readonly PictureBox pic;
        readonly Label lbl;

        public PreviewForm(Session session, List<PageRef> pages, int index, string title)
        {
            this.session = session;
            this.pages = pages;
            this.index = index;
            float s;
            using (var g = CreateGraphics()) s = g.DpiX / 96f;
            Text = "Xem trước - " + title;
            Font = new Font("Segoe UI", 9f);
            var wa = Screen.FromPoint(Cursor.Position).WorkingArea;
            Size = new Size(Math.Min(wa.Width, (int)(wa.Height * 0.8)), (int)(wa.Height * 0.9));
            StartPosition = FormStartPosition.CenterParent;
            KeyPreview = true;
            ShowInTaskbar = false;

            pic = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.FromArgb(50, 50, 50) };
            var bar = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding((int)(4 * s)) };
            var prev = new Button { Text = "◀ Trang trước", AutoSize = true };
            var next = new Button { Text = "Trang sau ▶", AutoSize = true };
            lbl = new Label { AutoSize = true, Margin = new Padding((int)(10 * s), (int)(8 * s), 0, 0) };
            prev.Click += (o, e) => Go(-1);
            next.Click += (o, e) => Go(1);
            bar.Controls.AddRange(new Control[] { prev, next, lbl });
            Controls.Add(pic);
            Controls.Add(bar);
            KeyDown += (o, e) =>
            {
                if (e.KeyCode == Keys.Left || e.KeyCode == Keys.PageUp) Go(-1);
                else if (e.KeyCode == Keys.Right || e.KeyCode == Keys.PageDown) Go(1);
                else if (e.KeyCode == Keys.Escape) Close();
            };
            Shown += (o, e) => ShowPage();
            FormClosed += (o, e) => { if (pic.Image != null) pic.Image.Dispose(); };
        }

        void Go(int d)
        {
            int n = index + d;
            if (n < 0 || n >= pages.Count) return;
            index = n;
            ShowPage();
        }

        void ShowPage()
        {
            lbl.Text = "Trang " + (index + 1) + " / " + pages.Count;
            Cursor = Cursors.WaitCursor;
            try
            {
                var old = pic.Image;
                pic.Image = session.Render(pages[index], 110);
                if (old != null) old.Dispose();
            }
            catch (Exception ex) { lbl.Text += "   (" + ex.Message + ")"; }
            finally { Cursor = Cursors.Default; }
        }
    }

    // ------------------------------------------------------------------
    // Thẻ "Soạn PDF" - tương tự PDF24 Creator
    // ------------------------------------------------------------------
    class CreatorView : UserControl
    {
        readonly float scale;
        int S(int v) { return (int)(v * scale); }
        readonly Session session = new Session();
        readonly FileBrowser browser;
        readonly Workspace ws;
        readonly SplitContainer split;
        readonly ToolStripStatusLabel status;
        readonly StatusStrip statusStrip;
        readonly List<ToolStripItem> needSelection = new List<ToolStripItem>(), needDoc = new List<ToolStripItem>();
        ToolStripItem bMerge, bCloseAll;
        public Action<string> SendToToolbox;

        public CreatorView(float scale)
        {
            this.scale = scale;
            Font = new Font("Segoe UI", 9f);

            split = new SplitContainer { Dock = DockStyle.Fill, FixedPanel = FixedPanel.Panel1, SplitterWidth = S(5) };
            browser = new FileBrowser(scale) { Dock = DockStyle.Fill };
            browser.OpenFiles = files => OpenFiles(files, null, -1);
            split.Panel1.Controls.Add(browser);

            ws = new Workspace(scale) { Dock = DockStyle.Fill, Session = session };
            ws.SaveRequested = d => SaveDoc(d);
            ws.CloseRequested = d => CloseDoc(d);
            ws.FilesDropped = (f, d, i) => OpenFiles(f, d, i);
            ws.PreviewRequested = Preview;
            ws.Changed += UpdateUi;
            session.ThumbsChanged = () =>
            {
                try { if (ws.IsHandleCreated) ws.BeginInvoke(new Action(ws.Invalidate)); } catch { }
            };

            var bar = new ToolStrip { GripStyle = ToolStripGripStyle.Hidden, Padding = new Padding(S(4), S(2), S(4), S(2)), ImageScalingSize = new Size(S(16), S(16)) };
            Func<string, string, Action, ToolStripButton> add = (text, tip, act) =>
            {
                var b = new ToolStripButton(text) { ToolTipText = tip, DisplayStyle = ToolStripItemDisplayStyle.Text, Padding = new Padding(S(3), S(2), S(3), S(2)) };
                b.Click += (o, e) => act();
                bar.Items.Add(b);
                return b;
            };
            add("Mở...", "Mở file PDF hoặc ảnh (Ctrl+O)", OpenDialog);
            needDoc.Add(add("Lưu", "Lưu tài liệu đang chọn (Ctrl+S)", () => SaveDoc(ws.Active)));
            bar.Items.Add(new ToolStripSeparator());
            bMerge = add("Gộp tất cả", "Gộp tất cả tài liệu thành một", ws.MergeAll);
            needSelection.Add(add("Tách trang chọn", "Chuyển các trang đang chọn thành tài liệu mới", ws.ExtractSelected));
            bar.Items.Add(new ToolStripSeparator());
            needSelection.Add(add("↺ Xoay trái", "Xoay trang đang chọn 90° ngược chiều kim đồng hồ", () => ws.RotateSelected(270)));
            needSelection.Add(add("↻ Xoay phải", "Xoay trang đang chọn 90° theo chiều kim đồng hồ", () => ws.RotateSelected(90)));
            needSelection.Add(add("Xóa trang", "Xóa các trang đang chọn (Delete)", ws.DeleteSelected));
            needSelection.Add(add("Xem trước", "Xem trang đang chọn cỡ lớn (Enter / nhấn đúp)", () => { var s = ws.OrderedSelection(); if (s.Count > 0) Preview(s[0]); }));
            bar.Items.Add(new ToolStripSeparator());
            add("＋", "Phóng to ảnh trang (Ctrl + lăn chuột)", () => ws.Zoom(S(30)));
            add("－", "Thu nhỏ ảnh trang (Ctrl + lăn chuột)", () => ws.Zoom(-S(30)));
            bar.Items.Add(new ToolStripSeparator());
            needDoc.Add(add("Nén / OCR / khác...", "Mở tài liệu đang chọn trong Hộp công cụ (nén, OCR, đánh số trang, mật khẩu...)", SendActiveToToolbox));
            needDoc.Add(add("Đóng", "Đóng tài liệu đang chọn", () => CloseDoc(ws.Active)));
            bCloseAll = add("Đóng tất cả", "Đóng tất cả tài liệu", CloseAll);

            ws.PageMenu = new ContextMenuStrip();
            ws.PageMenu.Items.Add("Xem trước", null, (o, e) => { var s = ws.OrderedSelection(); if (s.Count > 0) Preview(s[0]); });
            ws.PageMenu.Items.Add("Xoay trái", null, (o, e) => ws.RotateSelected(270));
            ws.PageMenu.Items.Add("Xoay phải", null, (o, e) => ws.RotateSelected(90));
            ws.PageMenu.Items.Add("Tách thành tài liệu mới", null, (o, e) => ws.ExtractSelected());
            ws.PageMenu.Items.Add("Xóa", null, (o, e) => ws.DeleteSelected());
            ws.DocMenu = new ContextMenuStrip();
            ws.DocMenu.Items.Add("Lưu...", null, (o, e) => SaveDoc(ws.Active));
            ws.DocMenu.Items.Add("Đổi tên...", null, (o, e) => RenameActive());
            ws.DocMenu.Items.Add("Chọn tất cả trang", null, (o, e) => ws.SelectAll());
            ws.DocMenu.Items.Add("Mở trong Hộp công cụ...", null, (o, e) => SendActiveToToolbox());
            ws.DocMenu.Items.Add("Đóng", null, (o, e) => CloseDoc(ws.Active));

            statusStrip = new StatusStrip();
            status = new ToolStripStatusLabel { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
            statusStrip.Items.Add(status);

            split.Panel2.Controls.Add(ws);
            split.Panel2.Controls.Add(bar);
            Controls.Add(split);
            Controls.Add(statusStrip);

            Load += (o, e) =>
            {
                try { split.SplitterDistance = S(int.Parse(Settings.Get("browserwidth", "700"))); } catch { }
                browser.Navigate(Settings.Get("browser", Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)), false);
                UpdateUi();
            };
        }

        public bool HasUnsaved { get { return ws.Docs.Any(d => d.Modified); } }

        public void SaveSettings()
        {
            Settings.Set("browserwidth", ((int)(split.SplitterDistance / scale)).ToString());
        }

        public void Cleanup() { session.Dispose(); }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.S)) { SaveDoc(ws.Active); return true; }
            if (keyData == (Keys.Control | Keys.O)) { OpenDialog(); return true; }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        void UpdateUi()
        {
            bool sel = ws.Selected.Count > 0, doc = ws.Active != null;
            foreach (var b in needSelection) b.Enabled = sel;
            foreach (var b in needDoc) b.Enabled = doc;
            bMerge.Enabled = ws.Docs.Count > 1;
            bCloseAll.Enabled = ws.Docs.Count > 0;
            int pages = ws.Docs.Sum(d => d.Pages.Count);
            status.Text = ws.Docs.Count == 0
                ? "Kéo file từ danh sách bên trái hoặc từ Explorer vào vùng bên phải. Kéo trang để sắp xếp, giữ Ctrl khi thả để sao chép."
                : ws.Docs.Count + " tài liệu, " + pages + " trang" + (sel ? ", đang chọn " + ws.Selected.Count + " trang" : "") +
                  "   |   Kéo trang để sắp xếp hoặc chuyển sang tài liệu khác, giữ Ctrl để sao chép.";
        }

        void OpenDialog()
        {
            using (var d = new OpenFileDialog { Multiselect = true })
            {
                d.Filter = "PDF và ảnh|*.pdf;*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff|File PDF|*.pdf|Ảnh|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tif;*.tiff";
                if (browser.CurrentFolder != null) d.InitialDirectory = browser.CurrentFolder;
                if (d.ShowDialog(FindForm()) == DialogResult.OK) OpenFiles(d.FileNames, null, -1);
            }
        }

        public void OpenFiles(IEnumerable<string> paths, PdfDoc target, int index)
        {
            var files = new List<string>();
            foreach (string p in paths)
            {
                if (Directory.Exists(p))
                    files.AddRange(Directory.GetFiles(p).Where(FileBrowser.Supported).OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase));
                else if (File.Exists(p) && FileBrowser.Supported(p)) files.Add(p);
            }
            if (files.Count == 0) { status.Text = "Chỉ mở được file PDF hoặc ảnh (JPG, PNG, BMP, GIF, TIFF)."; return; }
            if (index < 0 && target != null) index = target.Pages.Count;
            Cursor = Cursors.WaitCursor;
            try
            {
                foreach (string f in files)
                {
                    status.Text = "Đang mở " + Path.GetFileName(f) + "...";
                    statusStrip.Refresh();
                    try
                    {
                        PdfDoc d = session.Import(f, FindForm());
                        if (d != null) ws.AddDoc(d, target, ref index);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(FindForm(), "Không mở được \"" + Path.GetFileName(f) + "\":\r\n" + ex.Message, "PdfTool", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
            finally { Cursor = Cursors.Default; }
            ws.Relayout();
            UpdateUi();
        }

        bool SaveDoc(PdfDoc d)
        {
            if (d == null) return false;
            using (var dlg = new SaveFileDialog { Filter = "File PDF (*.pdf)|*.pdf", DefaultExt = "pdf", FileName = d.Name + ".pdf", OverwritePrompt = true, Title = "Lưu tài liệu" })
            {
                string dir = d.FilePath != null ? Path.GetDirectoryName(d.FilePath) : browser.CurrentFolder;
                if (dir != null && Directory.Exists(dir)) dlg.InitialDirectory = dir;
                if (dlg.ShowDialog(FindForm()) != DialogResult.OK) return false;
                Cursor = Cursors.WaitCursor;
                try
                {
                    session.Save(d, dlg.FileName);
                    d.FilePath = dlg.FileName;
                    d.Name = Path.GetFileNameWithoutExtension(dlg.FileName);
                    d.Modified = false;
                    ws.Invalidate();
                    UpdateUi();
                    status.Text = "Đã lưu: " + dlg.FileName;
                    return true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(FindForm(), "Không lưu được:\r\n" + ex.Message, "PdfTool", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                finally { Cursor = Cursors.Default; }
            }
        }

        // false = người dùng hủy
        bool ConfirmClose(PdfDoc d)
        {
            if (!d.Modified) return true;
            var r = MessageBox.Show(FindForm(), "Tài liệu \"" + d.Name + "\" có thay đổi chưa lưu. Lưu trước khi đóng?", "PdfTool",
                                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (r == DialogResult.Cancel) return false;
            return r == DialogResult.No || SaveDoc(d);
        }

        void CloseDoc(PdfDoc d)
        {
            if (d == null || !ConfirmClose(d)) return;
            ws.RemoveDoc(d);
        }

        void CloseAll()
        {
            if (HasUnsaved && MessageBox.Show(FindForm(), "Có tài liệu chưa lưu. Vẫn đóng tất cả?", "PdfTool", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            ws.Docs.Clear();
            ws.Relayout();
            UpdateUi();
        }

        void RenameActive()
        {
            var d = ws.Active;
            if (d == null) return;
            string name = InputDialog.Show(FindForm(), "Đổi tên", "Tên tài liệu:", d.Name, false);
            if (string.IsNullOrWhiteSpace(name)) return;
            foreach (char c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
            d.Name = name.Trim();
            ws.Invalidate();
        }

        void Preview(PageRef p)
        {
            var d = ws.Docs.FirstOrDefault(x => x.Pages.Contains(p));
            if (d == null) return;
            using (var f = new PreviewForm(session, d.Pages.ToList(), d.Pages.IndexOf(p), d.Name)) f.ShowDialog(FindForm());
        }

        void SendActiveToToolbox()
        {
            var d = ws.Active;
            if (d == null || SendToToolbox == null) return;
            if (d.Modified || d.FilePath == null)
            {
                if (MessageBox.Show(FindForm(), "Cần lưu tài liệu \"" + d.Name + "\" trước khi dùng Hộp công cụ. Lưu ngay?", "PdfTool",
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes || !SaveDoc(d)) return;
            }
            SendToToolbox(d.FilePath);
        }
    }

    // ------------------------------------------------------------------
    // Cửa sổ chính
    // ------------------------------------------------------------------
    class MainForm : Form
    {
        readonly float scale = 1f;
        int S(int v) { return (int)(v * scale); }
        readonly TabControl tabs;
        readonly CreatorView creator;
        readonly ToolboxView toolbox;

        public MainForm(string[] args)
        {
            using (var g = CreateGraphics()) scale = g.DpiX / 96f;
            Text = "PdfTool - Bộ công cụ PDF";
            Font = new Font("Segoe UI", 9f);
            Size = new Size(S(1400), S(860));
            MinimumSize = new Size(S(1000), S(640));
            StartPosition = FormStartPosition.CenterScreen;
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }
            if (Settings.Get("maximized", "1") == "1") WindowState = FormWindowState.Maximized;

            tabs = new TabControl { Dock = DockStyle.Fill, Padding = new Point(S(18), S(5)) };
            creator = new CreatorView(scale) { Dock = DockStyle.Fill };
            toolbox = new ToolboxView(scale) { Dock = DockStyle.Fill };
            var t1 = new TabPage("Soạn PDF");
            t1.Controls.Add(creator);
            var t2 = new TabPage("Hộp công cụ");
            t2.Controls.Add(toolbox);
            tabs.TabPages.Add(t1);
            tabs.TabPages.Add(t2);
            Controls.Add(tabs);
            if (Settings.Get("tab", "0") == "1") tabs.SelectedIndex = 1;

            creator.SendToToolbox = path => { tabs.SelectedIndex = 1; toolbox.AddPdf(path); };
            Load += (s, e) =>
            {
                if (args.Length > 0) { tabs.SelectedIndex = 0; creator.OpenFiles(args, null, -1); }
            };
            FormClosing += (s, e) =>
            {
                if (toolbox.IsBusy)
                {
                    if (MessageBox.Show(this, "Hộp công cụ đang xử lý. Dừng và thoát?", "PdfTool", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) { e.Cancel = true; return; }
                    toolbox.CancelRunning();
                }
                if (creator.HasUnsaved &&
                    MessageBox.Show(this, "Có tài liệu chưa lưu trong thẻ Soạn PDF. Vẫn thoát?", "PdfTool", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                { e.Cancel = true; return; }
                toolbox.SaveSettings();
                creator.SaveSettings();
                Settings.Set("tab", tabs.SelectedIndex.ToString());
                Settings.Set("maximized", WindowState == FormWindowState.Maximized ? "1" : "0");
                Settings.Save();
                creator.Cleanup();
            };
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
            Application.ThreadException += (s, e) =>
                MessageBox.Show(e.Exception.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            Settings.Load();
            Application.Run(new MainForm(args));
        }
    }
}
