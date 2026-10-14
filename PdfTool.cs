// PdfTool - bộ công cụ PDF (gộp, tách, nén, xoay, chuyển đổi, OCR, đánh số trang, đóng dấu, mật khẩu)
// Dùng các công cụ dòng lệnh mã nguồn mở: qpdf, Ghostscript, Tesseract (tìm trong thư mục tools\ cạnh exe,
// hoặc trong thư mục cài đặt PDF24, hoặc trong PATH).
// Chuyển đổi Word, Excel, PowerPoint dùng Microsoft Office có sẵn trên máy (điều khiển qua COM).
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
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

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
        static readonly Dictionary<Thread, List<Process>> running = new Dictionary<Thread, List<Process>>();

        public static void Track(Process p)
        {
            lock (running)
            {
                List<Process> list;
                if (!running.TryGetValue(Thread.CurrentThread, out list)) running[Thread.CurrentThread] = list = new List<Process>();
                list.Add(p);
            }
        }

        public static void Untrack(Process p)
        {
            lock (running)
            {
                List<Process> list;
                if (running.TryGetValue(Thread.CurrentThread, out list) && list.Remove(p) && list.Count == 0) running.Remove(Thread.CurrentThread);
            }
        }

        public static void Kill(Thread t)
        {
            List<Process> list;
            lock (running)
            {
                if (!running.TryGetValue(t, out list)) return;
                list = new List<Process>(list);
            }
            foreach (Process p in list)
                try { if (!p.HasExited) p.Kill(); } catch { }
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
                Track(p);
                p.BeginOutputReadLine();
                p.BeginErrorReadLine();
                p.WaitForExit();
                Untrack(p);
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

        // ---------------- Microsoft Office ----------------
        // Office chạy ẩn nên không ai trả lời được hộp thoại hỏi mật khẩu: đưa mật khẩu sai để file có mật khẩu báo lỗi ngay
        const string WrongPassword = "PdfTool-khong-co-mat-khau";

        // File Office kiểu mới (docx, xlsx, pptx...) là file zip; khi đặt mật khẩu mở, Office lưu thành file OLE mã hóa.
        // PowerPoint không nhận mật khẩu qua COM nên phải phát hiện trước khi mở.
        static void EnsureNoPassword(string f, string app)
        {
            if (Path.GetExtension(f).Length != 5) return; // .docx, .xlsm, .pptx... (đuôi kiểu cũ .doc, .xls, .ppt chỉ có 3 chữ)
            var head = new byte[4];
            using (var s = File.OpenRead(f)) s.Read(head, 0, 4);
            if (head[0] == 0xD0 && head[1] == 0xCF && head[2] == 0x11 && head[3] == 0xE0)
                throw new InvalidOperationException("File được đặt mật khẩu mở. Hãy mở bằng " + app + ", gỡ mật khẩu rồi thử lại.");
        }

        public static void WordToPdf(Job job, bool forScreen, bool bookmarks)
        {
            using (var word = OfficeApp.Word(job))
            {
                object docs = word.Get(word.App, "Documents");
                job.Each(f =>
                {
                    EnsureNoPassword(f, "Word");
                    string outPath = job.Out(f, "", ".pdf");
                    object doc = word.Call(docs, "Open", "FileName", f, "ConfirmConversions", false, "ReadOnly", true, "AddToRecentFiles", false,
                                           "PasswordDocument", WrongPassword, "Visible", false, "NoEncodingDialog", true);
                    try
                    {
                        word.Call(doc, "ExportAsFixedFormat", "OutputFileName", outPath, "ExportFormat", 17 /* wdExportFormatPDF */, "OpenAfterExport", false,
                                  "OptimizeFor", forScreen ? 1 : 0, "IncludeDocProps", true, "CreateBookmarks", bookmarks ? 1 : 0 /* dấu trang theo đề mục */,
                                  "DocStructureTags", true);
                    }
                    finally { word.Close(doc, "SaveChanges", 0); }
                    job.Done(outPath);
                });
            }
        }

        public static void ExcelToPdf(Job job, bool activeSheetOnly, bool fitWidth)
        {
            using (var excel = OfficeApp.Excel(job))
            {
                object books = excel.Get(excel.App, "Workbooks");
                job.Each(f =>
                {
                    EnsureNoPassword(f, "Excel");
                    string outPath = job.Out(f, "", ".pdf");
                    object wb = excel.Call(books, "Open", "Filename", f, "UpdateLinks", 0, "ReadOnly", true, "Password", WrongPassword,
                                           "IgnoreReadOnlyRecommended", true, "Notify", false, "AddToMru", false);
                    try
                    {
                        if (fitWidth) FitToPageWidth(job, excel, wb, activeSheetOnly);
                        excel.Call(activeSheetOnly ? excel.Get(wb, "ActiveSheet") : wb, "ExportAsFixedFormat", "Type", 0 /* xlTypePDF */, "Filename", outPath,
                                   "Quality", 0, "IncludeDocProperties", true, "IgnorePrintAreas", false, "OpenAfterPublish", false);
                    }
                    finally { excel.Close(wb, "SaveChanges", false); }
                    job.Done(outPath);
                });
            }
        }

        // Co mỗi trang tính vừa một trang giấy theo chiều ngang (chiều dọc bao nhiêu trang tùy nội dung). Chỉ đổi bản đang mở, không lưu vào file.
        static void FitToPageWidth(Job job, OfficeApp excel, object wb, bool activeSheetOnly)
        {
            try
            {
                var sheets = new List<object>();
                if (activeSheetOnly) sheets.Add(excel.Get(wb, "ActiveSheet"));
                else
                {
                    object all = excel.Get(wb, "Worksheets");
                    int n = (int)excel.Get(all, "Count");
                    for (int i = 1; i <= n; i++) sheets.Add(excel.Get(all, "Item", i));
                }
                excel.Set(excel.App, "PrintCommunication", false); // gom các thay đổi trang in, nhanh hơn nhiều khi có nhiều trang tính
                try
                {
                    foreach (object sh in sheets)
                    {
                        object ps = excel.Get(sh, "PageSetup");
                        excel.Set(ps, "Zoom", false);
                        excel.Set(ps, "FitToPagesWide", 1);
                        excel.Set(ps, "FitToPagesTall", false);
                    }
                }
                finally { excel.Set(excel.App, "PrintCommunication", true); }
            }
            catch (Exception ex)
            {
                if (job.Cancelled) throw;
                job.Log("   Không co được trang tính (" + ex.Message + "). Giữ nguyên thiết lập trang in của file.");
            }
        }

        public static void PowerPointToPdf(Job job)
        {
            using (var ppt = OfficeApp.PowerPoint(job))
            {
                object list = ppt.Get(ppt.App, "Presentations");
                job.Each(f =>
                {
                    EnsureNoPassword(f, "PowerPoint");
                    string outPath = job.Out(f, "", ".pdf");
                    object pres = ppt.Call(list, "Open", "FileName", f, "ReadOnly", -1 /* msoTrue */, "WithWindow", 0 /* msoFalse: không hiện cửa sổ */);
                    try { ppt.Call(pres, "SaveAs", "FileName", outPath, "FileFormat", 32 /* ppSaveAsPDF */); }
                    finally { ppt.Close(pres); }
                    job.Done(outPath);
                });
            }
        }

        public static void PdfToWord(Job job)
        {
            // Mỗi lần mở PDF, Word hỏi "Word sẽ chuyển PDF thành tài liệu có thể chỉnh sửa..." và chờ bấm OK; Word chạy ẩn nên sẽ treo.
            // Tạm bật tùy chọn "không hiện lại thông báo này" của Word (đặt trước khi mở Word), xong việc trả lại như cũ.
            using (new TempRegValue(OfficeApp.RegKey("Word") + @"\Options", "DisableConvertPdfWarning", 1))
            using (var word = OfficeApp.Word(job))
            {
                if (word.Version < 15) throw new InvalidOperationException("Cần Microsoft Word 2013 trở lên để chuyển PDF sang Word.");
                object docs = word.Get(word.App, "Documents");
                job.Each(f =>
                {
                    Ext.EnsureOpenable(f);
                    string outPath = job.Out(f, "", ".docx");
                    job.Log("   Word đang đọc PDF, file nhiều trang có thể mất vài phút...");
                    object doc = word.Call(docs, "Open", "FileName", f, "ConfirmConversions", false, "ReadOnly", true, "AddToRecentFiles", false, "Visible", false);
                    try { word.Call(doc, "SaveAs2", "FileName", outPath, "FileFormat", 16 /* wdFormatDocumentDefault: .docx */, "AddToRecentFiles", false); }
                    finally { word.Close(doc, "SaveChanges", 0); }
                    job.Done(outPath);
                });
            }
        }
    }

    // ------------------------------------------------------------------
    // Một phiên Microsoft Office (Word, Excel, PowerPoint) chạy ẩn, điều khiển qua COM.
    // Gọi theo tên (late binding) nên không cần thư viện Interop và dùng được với mọi phiên bản Office.
    // ------------------------------------------------------------------
    public class OfficeApp : IDisposable
    {
        public readonly object App;
        readonly Job job;
        readonly string name;
        readonly Process proc;                   // Office do PdfTool mở; null = dùng chung PowerPoint người dùng đang mở sẵn
        readonly List<object> refs = new List<object>();
        readonly List<KeyValuePair<string, object>> changed = new List<KeyValuePair<string, object>>();

        const int RPC_E_CALL_REJECTED = unchecked((int)0x80010001), RPC_E_SERVERCALL_RETRYLATER = unchecked((int)0x8001010A);

        public static OfficeApp Word(Job job) { return Start(job, "Word", "WINWORD", "DisplayAlerts", 0 /* wdAlertsNone */, "AutomationSecurity", 3 /* không chạy macro */); }
        public static OfficeApp Excel(Job job) { return Start(job, "Excel", "EXCEL", "DisplayAlerts", false, "AutomationSecurity", 3); }
        public static OfficeApp PowerPoint(Job job) { return Start(job, "PowerPoint", "POWERPNT", "DisplayAlerts", 1 /* ppAlertsNone */, "AutomationSecurity", 3); }

        // Mở Office rồi đặt các tùy chọn (cặp tên, giá trị) để không hiện hộp thoại nào
        static OfficeApp Start(Job job, string name, string exe, params object[] options)
        {
            var o = new OfficeApp(job, name, exe);
            try
            {
                for (int i = 0; i + 1 < options.Length; i += 2)
                {
                    string prop = (string)options[i];
                    o.changed.Add(new KeyValuePair<string, object>(prop, o.Get(o.App, prop)));
                    o.Set(o.App, prop, options[i + 1]);
                }
                return o;
            }
            catch { o.Dispose(); throw; }
        }

        OfficeApp(Job job, string name, string exe)
        {
            this.job = job;
            this.name = name;
            Type type = Type.GetTypeFromProgID(name + ".Application");
            if (type == null) throw new InvalidOperationException("Máy chưa cài Microsoft " + name + ". Công cụ này dùng " + name + " có sẵn trên máy để chuyển đổi.");
            ForgetCrashes(name, job.Inputs);
            job.Log("Đang khởi động Microsoft " + name + "...");
            int[] before = Pids(exe);
            try { App = Activator.CreateInstance(type); }
            catch (Exception ex) { throw new InvalidOperationException("Không khởi động được Microsoft " + name + ": " + ex.Message); }
            refs.Add(App);
            // Tiến trình vừa mở thêm: "Dừng" sẽ tắt nó. PowerPoint chỉ chạy một bản nên nếu đang mở sẵn thì dùng chung, không tắt.
            int pid = Pids(exe).Except(before).FirstOrDefault();
            if (pid != 0) try { proc = Process.GetProcessById(pid); } catch { }
            if (proc != null) Ext.Track(proc);
        }

        static int[] Pids(string exe)
        {
            Process[] ps = Process.GetProcessesByName(exe);
            int[] ids = ps.Select(p => p.Id).ToArray();
            foreach (Process p in ps) p.Dispose();
            return ids;
        }

        // Khóa cài đặt của ứng dụng trong HKEY_CURRENT_USER, ví dụ Software\Microsoft\Office\16.0\Word (đọc được khi chưa mở Office)
        public static string RegKey(string name)
        {
            string ver = null;
            using (RegistryKey k = Registry.ClassesRoot.OpenSubKey(name + @".Application\CurVer"))
                if (k != null) ver = k.GetValue("") as string;                        // "Word.Application.16"
            int dot = ver == null ? -1 : ver.LastIndexOf('.');
            return @"Software\Microsoft\Office\" + (dot > 0 ? ver.Substring(dot + 1) + ".0" : "16.0") + @"\" + name;
        }

        // Office bị tắt ngang khi đang mở một file (ví dụ bấm "Dừng") sẽ ghi nhớ file đó; lần sau mở lại thì hỏi "Lần trước mở file này
        // gây lỗi nghiêm trọng, vẫn muốn mở?" và vì chạy ẩn nên sẽ treo. Xóa ghi nhớ của đúng các file sắp chuyển đổi.
        static void ForgetCrashes(string name, List<string> files)
        {
            foreach (string sub in new[] { "DisabledItems", "StartupItems" })
                try
                {
                    using (RegistryKey k = Registry.CurrentUser.OpenSubKey(RegKey(name) + @"\Resiliency\" + sub, true))
                    {
                        if (k == null) continue;
                        foreach (string v in k.GetValueNames())
                        {
                            var data = k.GetValue(v) as byte[];
                            if (data == null || data.Length < 2) continue;
                            // Đường dẫn lưu dạng UTF-16 trong dữ liệu nhị phân; đọc cả hai cách căn byte
                            string text = Encoding.Unicode.GetString(data) + "\n" + Encoding.Unicode.GetString(data, 1, data.Length - 1);
                            if (files.Any(f => text.IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0)) k.DeleteValue(v, false);
                        }
                    }
                }
                catch { }
        }

        public double Version
        {
            get
            {
                double v;
                return double.TryParse(Get(App, "Version") as string, NumberStyles.Float, CultureInfo.InvariantCulture, out v) ? v : 0;
            }
        }

        public object Get(object target, string prop, params object[] index) { return Invoke(target, prop, BindingFlags.GetProperty, index, null); }

        public void Set(object target, string prop, object value) { Invoke(target, prop, BindingFlags.SetProperty, new[] { value }, null); }

        // Gọi phương thức với tham số theo tên: Call(doc, "Close", "SaveChanges", 0)
        public object Call(object target, string method, params object[] namedArgs)
        {
            int n = namedArgs.Length / 2;
            var names = new string[n];
            var values = new object[n];
            for (int i = 0; i < n; i++) { names[i] = (string)namedArgs[2 * i]; values[i] = namedArgs[2 * i + 1]; }
            return Invoke(target, method, BindingFlags.InvokeMethod, values, n > 0 ? names : null);
        }

        // Đóng tài liệu, bỏ qua lỗi (Office có thể đã bị dừng); tài liệu còn sót sẽ đóng theo khi thoát Office
        public void Close(object doc, params object[] namedArgs)
        {
            try { Call(doc, "Close", namedArgs); } catch { }
        }

        object Invoke(object target, string member, BindingFlags kind, object[] args, string[] names)
        {
            DateTime giveUp = DateTime.Now.AddSeconds(60);
            while (true)
            {
                try
                {
                    object r = target.GetType().InvokeMember(member, kind, null, target, args, null, null, names);
                    if (r != null && Marshal.IsComObject(r)) refs.Add(r);
                    return r;
                }
                catch (Exception ex)
                {
                    if (job.Cancelled) throw new OperationCanceledException();
                    Exception inner = ex is TargetInvocationException && ex.InnerException != null ? ex.InnerException : ex;
                    var com = inner as COMException;
                    int code = com != null ? com.ErrorCode : 0;
                    // Office đang bận (vừa khởi động, đang mở file...) thì từ chối lời gọi: chờ rồi gọi lại
                    if (code == RPC_E_CALL_REJECTED || code == RPC_E_SERVERCALL_RETRYLATER)
                    {
                        if (DateTime.Now < giveUp) { Thread.Sleep(250); continue; }
                        throw new InvalidOperationException("Microsoft " + name + " không phản hồi (có thể đang chờ trả lời một hộp thoại).");
                    }
                    if (code == unchecked((int)0x800706BA) || code == unchecked((int)0x800706BE) || code == unchecked((int)0x80010108))
                        throw new InvalidOperationException("Microsoft " + name + " đã thoát đột ngột.");
                    throw new InvalidOperationException(name + " báo lỗi: " + inner.Message.Trim(), inner);
                }
            }
        }

        // Thoát Office nếu do PdfTool mở và không ai đang dùng. Nếu người dùng mở file vào đúng phiên này
        // (cửa sổ đã hiện lên) hoặc đây là PowerPoint đang mở sẵn: để lại, chỉ trả lại các tùy chọn đã đổi.
        public void Dispose()
        {
            bool quit = proc != null && !IsVisible();
            try
            {
                if (quit) Call(App, "Quit");
                else for (int i = changed.Count - 1; i >= 0; i--) Set(App, changed[i].Key, changed[i].Value);
            }
            catch { }
            for (int i = refs.Count - 1; i >= 0; i--) try { Marshal.FinalReleaseComObject(refs[i]); } catch { }
            refs.Clear();
            if (proc == null) return;
            if (quit) try { if (!proc.WaitForExit(10000)) proc.Kill(); } catch { }
            Ext.Untrack(proc);
            proc.Dispose();
        }

        bool IsVisible()
        {
            try { return Convert.ToInt32(Get(App, "Visible")) != 0; }
            catch { return false; }
        }
    }

    // Đặt tạm một giá trị DWORD trong HKEY_CURRENT_USER; Dispose() trả lại giá trị cũ (hoặc xóa nếu trước đó chưa có)
    class TempRegValue : IDisposable
    {
        readonly string key, name;
        readonly object old;
        readonly RegistryValueKind oldKind;

        public TempRegValue(string key, string name, int value)
        {
            this.key = key;
            this.name = name;
            try
            {
                using (RegistryKey k = Registry.CurrentUser.CreateSubKey(key))
                {
                    old = k.GetValue(name);
                    if (old != null) oldKind = k.GetValueKind(name);
                    k.SetValue(name, value, RegistryValueKind.DWord);
                }
            }
            catch { }
        }

        public void Dispose()
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.CreateSubKey(key))
                    if (old == null) k.DeleteValue(name, false);
                    else k.SetValue(name, old, oldKind);
            }
            catch { }
        }
    }

    // ------------------------------------------------------------------
    // Định nghĩa công cụ và giao diện tùy chọn
    // ------------------------------------------------------------------
    // Bảng tùy chọn: cột nhãn bên trái, ô nhập giãn hết cột bên phải, gợi ý chữ nhỏ bên dưới ô
    class Opts : TableLayoutPanel
    {
        readonly List<Label> wrapping = new List<Label>();

        static int S(int v) { return Theme.S(v); }

        public Opts(float scale)
        {
            ColumnCount = 2;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Dock = DockStyle.Top;
            ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            Resize += (s, e) => { foreach (var l in wrapping) l.MaximumSize = new Size(Math.Max(S(100), ClientSize.Width - S(10)), 0); };
        }

        T Row<T>(string label, T c, Control layout, bool stretch, string hint) where T : Control
        {
            int r = RowCount++;
            RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, S(4), S(16), S(4)) }, 0, r);
            layout.Margin = new Padding(0, S(4), 0, S(4));
            layout.Anchor = stretch ? AnchorStyles.Left | AnchorStyles.Right : AnchorStyles.Left;
            Controls.Add(layout, 1, r);
            if (hint != null)
            {
                r = RowCount++;
                RowStyles.Add(new RowStyle(SizeType.AutoSize));
                var h = new Label { Text = hint, AutoSize = true, Font = Theme.Ui(9f), Margin = new Padding(S(2), 0, 0, S(6)) };
                Theme.Fg(h, TextRole.Muted);
                Controls.Add(h, 1, r);
            }
            return c;
        }

        public FSelect Combo(string label, int width, params string[] items)
        {
            var c = new FSelect();
            c.Items.AddRange(items);
            c.SelectedIndex = 0;
            return Row(label, c, c, true, null);
        }

        public TextBox Input(string label, int width, string value, string hint, bool password = false)
        {
            var t = new TextBox { Text = value, UseSystemPasswordChar = password };
            return Row(label, t, new FieldHost(t), true, hint);
        }

        public NumBox Num(string label, int min, int max, int value, string hint)
        {
            var n = new NumBox { Minimum = min, Maximum = max, Value = value };
            return Row(label, n, n, false, hint);
        }

        public CheckBox Check(string text, bool value)
        {
            int r = RowCount++;
            RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var c = new FCheck(text, value) { Margin = new Padding(0, S(4), 0, S(4)) };
            Controls.Add(c, 1, r);
            return c;
        }

        public ColorSwatch ColorPick(string label, Color value)
        {
            var b = new ColorSwatch(value);
            return Row(label, b, b, false, null);
        }

        public void Note(string text)
        {
            int r = RowCount++;
            RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var l = new Label { Text = text, AutoSize = true, Font = Theme.Ui(9f), MaximumSize = new Size(S(560), 0), Margin = new Padding(0, S(8), 0, S(4)) };
            Theme.Fg(l, TextRole.Muted);
            wrapping.Add(l);
            Controls.Add(l, 0, r);
            SetColumnSpan(l, 2);
        }
    }

    // Loại file đầu vào của công cụ
    class FileKind
    {
        public readonly string Name, Title, Glyph;   // tên dùng trong câu ("file Word"), tiêu đề danh sách file, biểu tượng ô kéo thả
        public readonly string[] Extensions;

        public FileKind(string name, string title, string glyph, params string[] extensions)
        {
            Name = name; Title = title; Glyph = glyph; Extensions = extensions;
        }

        // Bộ lọc của hộp thoại chọn file
        public string Filter
        {
            get
            {
                string p = string.Join(";", Extensions.Select(e => "*" + e));
                return char.ToUpperInvariant(Name[0]) + Name.Substring(1) + " (" + p + ")|" + p + "|Tất cả file|*.*";
            }
        }

        public static readonly FileKind Pdf = new FileKind("file PDF", "File PDF cần xử lý", "\uE896", ".pdf");
        public static readonly FileKind Image = new FileKind("ảnh", "Ảnh cần chuyển", "\uEB9F", ImageUtil.Extensions);
        public static readonly FileKind Word = new FileKind("file Word", "File Word cần chuyển", "\uE896", ".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm", ".rtf", ".odt");
        public static readonly FileKind Excel = new FileKind("file Excel", "File Excel cần chuyển", "\uE896", ".xls", ".xlsx", ".xlsm", ".xlsb", ".csv", ".ods");
        public static readonly FileKind PowerPoint = new FileKind("file PowerPoint", "File PowerPoint cần chuyển", "\uE896", ".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".pot", ".potx", ".odp");
    }

    class ToolDef
    {
        public string Name, Glyph, Group, Desc;
        public Color Color;
        public FileKind Input = FileKind.Pdf;
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
                var t = new ToolDef { Name = "Gộp PDF", Glyph = "\uEA3C", Group = "Trang", Color = Color.FromArgb(37, 99, 235), MinFiles = 2,
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
                var t = new ToolDef { Name = "Tách PDF", Glyph = "\uE8C6", Group = "Trang", Color = Color.FromArgb(22, 163, 74),
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
                var t = new ToolDef { Name = "Trích xuất / sắp xếp trang", Glyph = "\uE8CB", Group = "Trang", Color = Color.FromArgb(8, 145, 178),
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
                var t = new ToolDef { Name = "Xóa trang", Glyph = "\uE74D", Group = "Trang", Color = Color.FromArgb(220, 38, 38),
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
                var t = new ToolDef { Name = "Xoay trang", Glyph = "\uE7AD", Group = "Trang", Color = Color.FromArgb(124, 58, 237),
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
                var t = new ToolDef { Name = "Nén PDF", Glyph = "\uF012", Group = "Tối ưu & chuyển đổi", Color = Color.FromArgb(234, 88, 12),
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
                var t = new ToolDef { Name = "Ảnh sang PDF", Glyph = "\uEB9F", Group = "Tối ưu & chuyển đổi", Color = Color.FromArgb(219, 39, 119), Input = FileKind.Image,
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
                var t = new ToolDef { Name = "PDF sang ảnh", Glyph = "\uE8B9", Group = "Tối ưu & chuyển đổi", Color = Color.FromArgb(2, 132, 199),
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
                var t = new ToolDef { Name = "Nhận dạng chữ (OCR)", Glyph = "\uE8AC", Group = "Tối ưu & chuyển đổi", Color = Color.FromArgb(13, 148, 136),
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

            // --- Word sang PDF
            {
                var t = new ToolDef { Name = "Word sang PDF", Glyph = "W", Group = "Chuyển đổi Office", Color = Color.FromArgb(43, 87, 154), Input = FileKind.Word,
                    Desc = "Chuyển văn bản Word (DOC, DOCX, RTF, ODT) sang PDF, giữ nguyên bố cục, phông chữ và liên kết. Cần Microsoft Word trên máy." };
                var o = t.Options = new Opts(scale);
                var quality = o.Combo("Chất lượng", 260, "In ấn (chất lượng cao)", "Xem trên màn hình (file nhỏ hơn)");
                var marks = o.Check("Tạo dấu trang (bookmark) từ các đề mục", true);
                t.Prepare = () =>
                {
                    bool screen = quality.SelectedIndex == 1, b = marks.Checked;
                    return job => Ops.WordToPdf(job, screen, b);
                };
                tools.Add(t);
            }

            // --- Excel sang PDF
            {
                var t = new ToolDef { Name = "Excel sang PDF", Glyph = "X", Group = "Chuyển đổi Office", Color = Color.FromArgb(33, 115, 70), Input = FileKind.Excel,
                    Desc = "Chuyển bảng tính Excel (XLS, XLSX, CSV, ODS) sang PDF theo thiết lập trang in trong file. Cần Microsoft Excel trên máy." };
                var o = t.Options = new Opts(scale);
                var scope = o.Combo("Phạm vi", 260, "Tất cả trang tính", "Chỉ trang tính đang chọn khi lưu file");
                var fit = o.Check("Co mỗi trang tính vừa chiều ngang trang giấy", false);
                o.Note("Trang tính bị ẩn không được xuất. Vùng in (Print Area) đã đặt trong file vẫn được giữ. File gốc không bị thay đổi.");
                t.Prepare = () =>
                {
                    bool active = scope.SelectedIndex == 1, f = fit.Checked;
                    return job => Ops.ExcelToPdf(job, active, f);
                };
                tools.Add(t);
            }

            // --- PowerPoint sang PDF
            {
                var t = new ToolDef { Name = "PowerPoint sang PDF", Glyph = "P", Group = "Chuyển đổi Office", Color = Color.FromArgb(196, 62, 28), Input = FileKind.PowerPoint,
                    Desc = "Chuyển bài trình chiếu PowerPoint (PPT, PPTX, PPS, ODP) sang PDF, mỗi trang chiếu thành một trang. Cần Microsoft PowerPoint trên máy." };
                var o = t.Options = new Opts(scale);
                o.Note("Các bài trình chiếu bạn đang mở trong PowerPoint không bị ảnh hưởng.");
                t.Prepare = () => job => Ops.PowerPointToPdf(job);
                tools.Add(t);
            }

            // --- PDF sang Word
            {
                var t = new ToolDef { Name = "PDF sang Word", Glyph = "\uE932", Group = "Chuyển đổi Office", Color = Color.FromArgb(24, 90, 189),
                    Desc = "Chuyển PDF thành văn bản Word (DOCX) có thể chỉnh sửa, dùng tính năng đọc PDF của Microsoft Word 2013 trở lên." };
                var o = t.Options = new Opts(scale);
                o.Note("Bố cục phức tạp (bảng, nhiều cột, chữ trên hình) có thể bị lệch so với bản gốc. PDF scan chỉ có hình ảnh nên Word không lấy ra được chữ.");
                t.Prepare = () => job => Ops.PdfToWord(job);
                tools.Add(t);
            }

            // --- Đánh số trang
            {
                var t = new ToolDef { Name = "Đánh số trang", Glyph = "123", Group = "Chỉnh sửa", Color = Color.FromArgb(101, 163, 13),
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
                var t = new ToolDef { Name = "Đóng dấu chữ (watermark)", Glyph = "\uE8D2", Group = "Chỉnh sửa", Color = Color.FromArgb(192, 38, 211),
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
                    Color c = color.Value;
                    int op = (int)opacity.Value;
                    return job => Ops.Watermark(job, s, sz, c, op, a);
                };
                tools.Add(t);
            }

            // --- Đặt mật khẩu
            {
                var t = new ToolDef { Name = "Đặt mật khẩu", Glyph = "\uE72E", Group = "Bảo mật", Color = Color.FromArgb(71, 85, 105),
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
                var t = new ToolDef { Name = "Gỡ mật khẩu", Glyph = "\uE785", Group = "Bảo mật", Color = Color.FromArgb(100, 116, 139),
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
        int hoverItem = -1, lastToolIndex = -1;

        ListBox lstTools;
        Label lblTitle, lblDesc;
        IconBadge badge;
        Card cardFiles;
        FileList lvFiles;
        ImageList fileIcons;
        DropHint dropHint;
        Panel pnlOptions;
        FRadio rbSame, rbOther;
        TextBox txtOut;
        FButton btnAdd, btnRemove, btnUp, btnDown, btnClear, btnBrowse, btnRun, btnCancel, btnOpen;
        TextBox txtLog;
        ThinProgress progress;

        static int S(int v) { return Theme.S(v); }

        public ToolboxView(float scale)
        {
            Font = Theme.Ui(10f);
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

        static FButton Btn(string text, string glyph, ButtonKind kind, string tip, EventHandler click)
        {
            var b = new FButton(text, glyph, kind);
            if (tip != null) Theme.Tip(b, tip);
            b.Click += click;
            return b;
        }

        void BuildUi()
        {
            Theme.Bg(this, Role.Window);

            // --- Thanh bên: danh sách công cụ theo nhóm
            lstTools = new ListBox { Dock = DockStyle.Fill, DrawMode = DrawMode.OwnerDrawVariable, BorderStyle = BorderStyle.None, IntegralHeight = false };
            Theme.Bg(lstTools, Role.Surface);
            Theme.Native(lstTools);
            string group = null;
            foreach (var t in tools)
            {
                if (t.Group != group) { group = t.Group; lstTools.Items.Add(group); }
                lstTools.Items.Add(t);
            }
            lstTools.MeasureItem += (s, e) => e.ItemHeight = lstTools.Items[e.Index] is string ? S(e.Index == 0 ? 28 : 40) : S(44);
            lstTools.DrawItem += DrawTool;
            lstTools.SelectedIndexChanged += (s, e) => OnToolSelected();
            lstTools.MouseMove += (s, e) =>
            {
                int i = lstTools.IndexFromPoint(e.Location);
                if (i == hoverItem) return;
                int old = hoverItem;
                hoverItem = i;
                InvalidateItem(old);
                InvalidateItem(i);
            };
            lstTools.MouseLeave += (s, e) => { int old = hoverItem; hoverItem = -1; InvalidateItem(old); };

            var sidebar = new Panel { Dock = DockStyle.Left, Width = S(272), Padding = new Padding(S(8), S(6), S(8), 0) };
            Theme.Bg(sidebar, Role.Surface);
            sidebar.Controls.Add(lstTools);
            sidebar.Controls.Add(new DonatePanel(1f));
            var sideLine = new Panel { Dock = DockStyle.Left, Width = 1 };
            Theme.Bg(sideLine, Role.Divider);

            // --- Nội dung
            // Vùng nội dung: cuộn được khi cửa sổ thấp hơn chiều cao tối thiểu
            scrollHost = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
            Theme.Bg(scrollHost, Role.Window);
            Theme.Native(scrollHost);
            content = new TableLayoutPanel { Location = Point.Empty, ColumnCount = 1, RowCount = 2, Padding = new Padding(S(18), S(16), S(18), S(14)) };
            Theme.Bg(content, Role.Window);
            scrollHost.Controls.Add(content);
            scrollHost.Resize += (s, e) => FitContent();
            content.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            content.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var head = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, AutoSize = true, Margin = new Padding(S(6), 0, S(6), S(10)) };
            head.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            head.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            badge = new IconBadge { Size = new Size(S(52), S(52)), Margin = new Padding(0, S(2), S(16), 0), Anchor = AnchorStyles.Left | AnchorStyles.Top };
            var titles = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false, Dock = DockStyle.Fill, Margin = new Padding(0) };
            lblTitle = new Label { AutoSize = true, Font = Theme.UiSemi(18f), Margin = new Padding(0) };
            lblDesc = new Label { AutoSize = true, Font = Theme.Ui(10f), Margin = new Padding(S(2), S(2), 0, 0) };
            Theme.Fg(lblDesc, TextRole.Muted);
            titles.Controls.Add(lblTitle);
            titles.Controls.Add(lblDesc);
            head.Controls.Add(badge, 0, 0);
            head.Controls.Add(titles, 1, 0);
            content.Resize += (s, e) => lblDesc.MaximumSize = new Size(Math.Max(S(200), content.ClientSize.Width - S(120)), 0);

            var body = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Margin = new Padding(0) };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 52));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 48));
            body.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            // Thẻ danh sách file
            cardFiles = new Card("Tệp cần xử lý", "") { Dock = DockStyle.Fill };
            var fileBar = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = true, Margin = new Padding(0), Padding = new Padding(0, 0, 0, S(8)) };
            btnAdd = Btn("Thêm file", "", ButtonKind.Standard, "Chọn file để thêm vào danh sách", (s, e) => BrowseFiles());
            btnRemove = Btn("Bỏ file", "", ButtonKind.Subtle, "Bỏ các file đang chọn khỏi danh sách (Delete)", (s, e) => RemoveSelected());
            btnUp = Btn(null, "", ButtonKind.Subtle, "Đưa lên trên", (s, e) => MoveSelected(-1));
            btnDown = Btn(null, "", ButtonKind.Subtle, "Đưa xuống dưới", (s, e) => MoveSelected(1));
            btnClear = Btn("Bỏ tất cả", "", ButtonKind.Subtle, "Bỏ tất cả file khỏi danh sách", (s, e) => { tool.Files.Clear(); RefreshFiles(); });
            fileBar.Controls.AddRange(new Control[] { btnAdd, btnRemove, btnUp, btnDown, btnClear });

            var listHost = new Panel { Dock = DockStyle.Fill };
            fileIcons = new ImageList { ColorDepth = ColorDepth.Depth32Bit, ImageSize = new Size(S(16), S(16)) };
            lvFiles = new FileList { Dock = DockStyle.Fill, AllowDrop = true, FillColumn = 2, ClickableHeader = false, Images = fileIcons };
            Theme.Bg(lvFiles, Role.Surface);
            lvFiles.Columns.Add(new FileList.Column { Text = "Tên file", Width = S(250) });
            lvFiles.Columns.Add(new FileList.Column { Text = "Kích thước", Width = S(90), Right = true });
            lvFiles.Columns.Add(new FileList.Column { Text = "Thư mục", Width = S(200) });
            DragEventHandler enter = (s, e) => e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) && running == null ? DragDropEffects.Copy : DragDropEffects.None;
            DragEventHandler drop = (s, e) => AddFiles((string[])e.Data.GetData(DataFormats.FileDrop));
            lvFiles.DragEnter += enter;
            lvFiles.DragDrop += drop;
            lvFiles.SelectionChanged += (s, e) => UpdateButtons();
            lvFiles.KeyDown += (s, e) => { if (e.KeyCode == Keys.Delete) RemoveSelected(); };
            dropHint = new DropHint { Dock = DockStyle.Fill, AllowDrop = true };
            dropHint.Click += (s, e) => { if (running == null) BrowseFiles(); };
            dropHint.DragEnter += enter;
            dropHint.DragDrop += drop;
            listHost.Controls.Add(dropHint);
            listHost.Controls.Add(lvFiles);
            cardFiles.Controls.Add(listHost);
            cardFiles.Controls.Add(fileBar);

            // Cột phải: tùy chọn, nơi lưu, nút thực hiện, nhật ký
            var right = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Margin = new Padding(0) };
            right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            right.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var cardOpt = new Card("Tùy chọn", "") { Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
            pnlOptions = new Panel { Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
            cardOpt.Controls.Add(pnlOptions);

            var cardOut = new Card("Lưu kết quả", "") { Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
            var outGrid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 2, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink };
            outGrid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            outGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            outGrid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            rbSame = new FRadio("Cùng thư mục với file gốc", true) { Margin = new Padding(0, 0, 0, S(8)) };
            rbOther = new FRadio("Thư mục khác:", false) { Anchor = AnchorStyles.Left, Margin = new Padding(0, 0, S(10), 0) };
            txtOut = new TextBox();
            var outField = new FieldHost(txtOut) { Anchor = AnchorStyles.Left | AnchorStyles.Right, Margin = new Padding(0), LeadingGlyph = "" };
            btnBrowse = Btn(null, "", ButtonKind.Standard, "Chọn thư mục lưu kết quả", (s, e) => BrowseOut());
            btnBrowse.Anchor = AnchorStyles.Left;
            btnBrowse.Margin = new Padding(S(6), 0, 0, 0);
            btnBrowse.Size = new Size(outField.Height, outField.Height);
            outGrid.Controls.Add(rbSame, 0, 0);
            outGrid.SetColumnSpan(rbSame, 3);
            outGrid.Controls.Add(rbOther, 0, 1);
            outGrid.Controls.Add(outField, 1, 1);
            outGrid.Controls.Add(btnBrowse, 2, 1);
            txtOut.TextChanged += (s, e) => { if (txtOut.Focused) rbOther.Checked = true; };
            cardOut.Controls.Add(outGrid);

            var actions = new WrapFlow { Dock = DockStyle.Fill, AutoSize = true, Margin = new Padding(S(6), S(8), S(6), S(8)) };
            btnRun = Btn("Thực hiện", "", ButtonKind.Primary, "Bắt đầu xử lý các file trong danh sách", (s, e) => RunTool());
            btnRun.Font = Theme.UiSemi(10.5f);
            btnRun.Size = new Size(Math.Max(S(160), btnRun.GetPreferredSize(Size.Empty).Width), S(40));
            btnCancel = Btn("Dừng", "", ButtonKind.Standard, "Dừng công việc đang chạy", (s, e) => CancelRunning());
            btnCancel.Height = S(40);
            btnOpen = Btn("Mở thư mục kết quả", "", ButtonKind.Standard, "Mở thư mục chứa file vừa tạo", (s, e) => OpenResult());
            btnOpen.Height = S(40);
            btnRun.Margin = btnCancel.Margin = btnOpen.Margin = new Padding(0, 0, S(8), S(6));
            progress = new ThinProgress { Visible = false, Margin = new Padding(S(10), S(18), 0, 0) };
            actions.Controls.AddRange(new Control[] { btnRun, btnCancel, btnOpen, progress });

            var cardLog = new Card("Nhật ký", "") { Dock = DockStyle.Fill };
            txtLog = new TextBox
            {
                Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
                BorderStyle = BorderStyle.None, Font = Theme.Mono(9.5f)
            };
            Theme.Bg(txtLog, Role.Surface);
            cardLog.Controls.Add(txtLog);

            right.Controls.Add(cardOpt, 0, 0);
            right.Controls.Add(cardOut, 0, 1);
            right.Controls.Add(actions, 0, 2);
            right.Controls.Add(cardLog, 0, 3);
            body.Controls.Add(cardFiles, 0, 0);
            body.Controls.Add(right, 1, 0);
            content.Controls.Add(head, 0, 0);
            content.Controls.Add(body, 0, 1);

            this.head = head;
            this.right = right;
            this.cardOpt = cardOpt;
            this.cardOut = cardOut;
            this.actions = actions;
            Controls.Add(scrollHost);
            Controls.Add(sideLine);
            Controls.Add(sidebar);
            lstTools.SelectedItem = tools[0];
        }

        Panel scrollHost;
        TableLayoutPanel content, head, right;
        Card cardOpt, cardOut;
        FlowLayoutPanel actions;

        // Chiều cao tối thiểu = tiêu đề + các thẻ bên phải + chỗ cho nhật ký; thấp hơn thì hiện thanh cuộn
        void FitContent()
        {
            if (content == null || scrollHost.Width <= 0) return;
            int fullW = scrollHost.Width, fullH = scrollHost.Height;
            int rightW = Math.Max(S(200), (int)((fullW - content.Padding.Horizontal) * 0.48));
            int need = content.Padding.Vertical + head.GetPreferredSize(new Size(fullW, 0)).Height + head.Margin.Vertical +
                       cardOpt.GetPreferredSize(new Size(rightW, 0)).Height + cardOpt.Margin.Vertical +
                       cardOut.GetPreferredSize(new Size(rightW, 0)).Height + cardOut.Margin.Vertical +
                       actions.GetPreferredSize(new Size(rightW, 0)).Height + actions.Margin.Vertical + S(170);
            bool scroll = need > fullH;
            int w = fullW - (scroll ? SystemInformation.VerticalScrollBarWidth : 0);
            scrollHost.AutoScrollMinSize = scroll ? new Size(0, need) : Size.Empty;
            content.SetBounds(scrollHost.AutoScrollPosition.X, scrollHost.AutoScrollPosition.Y, w, scroll ? need : fullH);
        }

        void InvalidateItem(int i)
        {
            if (i >= 0 && i < lstTools.Items.Count) lstTools.Invalidate(lstTools.GetItemRectangle(i));
        }

        // Không cho chọn dòng tiêu đề nhóm; không đổi công cụ khi đang xử lý
        void OnToolSelected()
        {
            int i = lstTools.SelectedIndex;
            if (i < 0) return;
            if (lstTools.Items[i] is string)
            {
                int dir = i < lastToolIndex ? -1 : 1;
                for (int pass = 0; pass < 2; pass++, dir = -dir)
                    for (int j = i + dir; j >= 0 && j < lstTools.Items.Count; j += dir)
                        if (!(lstTools.Items[j] is string)) { lstTools.SelectedIndex = j; return; }
                return;
            }
            if (running != null && lastToolIndex >= 0 && i != lastToolIndex) { lstTools.SelectedIndex = lastToolIndex; return; }
            lastToolIndex = i;
            ShowTool((ToolDef)lstTools.Items[i]);
        }

        void DrawTool(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            var g = e.Graphics;
            var p = Theme.P;
            using (var b = new SolidBrush(p.Surface)) g.FillRectangle(b, e.Bounds);
            var header = lstTools.Items[e.Index] as string;
            if (header != null)
            {
                TextRenderer.DrawText(g, header.ToUpper(CultureInfo.CurrentCulture), Theme.UiSemi(8.5f),
                    new Rectangle(e.Bounds.X + S(12), e.Bounds.Bottom - S(24), e.Bounds.Width - S(12), S(20)), p.TextMuted, Theme.TextLeft);
                return;
            }
            var t = (ToolDef)lstTools.Items[e.Index];
            bool sel = (e.State & DrawItemState.Selected) != 0;
            var row = new RectangleF(e.Bounds.X + S(2), e.Bounds.Y + S(2), e.Bounds.Width - S(4), e.Bounds.Height - S(4));
            if (sel)
            {
                Theme.FillRound(g, row, S(6), p.Dark ? Color.FromArgb(255, 0x38, 0x38, 0x38) : Color.FromArgb(255, 0xEC, 0xEC, 0xEC), Color.Empty, 0);
                Theme.FillRound(g, new RectangleF(row.X, row.Y + row.Height / 2 - S(9), S(3), S(18)), Theme.Sf(1.5f), p.Accent, Color.Empty, 0);
            }
            else if (e.Index == hoverItem)
                Theme.FillRound(g, row, S(6), p.Hover, Color.Empty, 0);
            var icon = new RectangleF(row.X + S(12), row.Y + (row.Height - S(30)) / 2, S(30), S(30));
            Theme.Badge(g, icon, t.Color, t.Glyph, 11f);
            int x = (int)icon.Right + S(12);
            TextRenderer.DrawText(g, t.Name, sel ? Theme.UiSemi(10f) : Theme.Ui(10f), new Rectangle(x, e.Bounds.Y, e.Bounds.Right - x - S(6), e.Bounds.Height), p.Text, Theme.TextLeft);
        }

        // ------------------------------------------------------------------
        void ShowTool(ToolDef t)
        {
            if (t == null) return;
            tool = t;
            lblTitle.Text = t.Name;
            lblDesc.Text = t.Desc;
            badge.Fill = t.Color;
            badge.Glyph = t.Glyph;
            badge.Invalidate();
            pnlOptions.Controls.Clear();
            pnlOptions.Controls.Add(t.Options);
            Theme.Apply(t.Options);
            FitContent();
            cardFiles.Title = t.MinFiles > 1 ? "Các file PDF cần gộp" : t.Input.Title;
            dropHint.Title = "Kéo thả " + t.Input.Name + " vào đây";
            dropHint.Sub = "hoặc bấm vào đây để chọn file" + (t.MinFiles > 1 ? " (cần ít nhất 2 file)" : "");
            dropHint.Glyph = t.Input.Glyph;
            RefreshFiles();
        }

        bool Accepts(string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            return tool.Input.Extensions.Contains(ext);
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
                Log("Bỏ qua " + skipped + " file không phải " + tool.Input.Name + ".");
        }

        void BrowseFiles()
        {
            using (var d = new OpenFileDialog { Multiselect = true })
            {
                d.Filter = tool.Input.Filter;
                if (d.ShowDialog(this) == DialogResult.OK) AddFiles(d.FileNames);
            }
        }

        void RefreshFiles()
        {
            var sel = lvFiles.SelectedIndices.Select(i => tool.Files.ElementAtOrDefault(i)).ToList();
            lvFiles.BeginUpdate();
            lvFiles.Clear();
            foreach (string f in tool.Files)
            {
                string size = "";
                try { size = Job.SizeText(new FileInfo(f).Length); } catch { }
                lvFiles.Add(new[] { Path.GetFileName(f), size, Path.GetDirectoryName(f) }, ShellIcons.Key(fileIcons, f, false, false), f, sel.Contains(f));
            }
            lvFiles.EndUpdate();
            bool empty = tool.Files.Count == 0;
            dropHint.Visible = empty;
            lvFiles.Visible = !empty;
            dropHint.Invalidate();
            UpdateButtons();
        }

        void RemoveSelected()
        {
            foreach (int i in lvFiles.SelectedIndices.OrderByDescending(i => i)) tool.Files.RemoveAt(i);
            RefreshFiles();
        }

        void MoveSelected(int dir)
        {
            var si = lvFiles.SelectedIndices;
            if (si.Count != 1) return;
            int i = si[0], j = i + dir;
            if (j < 0 || j >= tool.Files.Count) return;
            string tmp = tool.Files[i]; tool.Files[i] = tool.Files[j]; tool.Files[j] = tmp;
            RefreshFiles();
            lvFiles.SelectOnly(j);
        }

        void UpdateButtons()
        {
            bool idle = running == null;
            var si = lvFiles.SelectedIndices;
            int nsel = si.Count;
            btnAdd.Enabled = btnBrowse.Enabled = idle;
            btnRemove.Enabled = idle && nsel > 0;
            btnUp.Enabled = idle && nsel == 1 && si[0] > 0;
            btnDown.Enabled = idle && nsel == 1 && si[0] < tool.Files.Count - 1;
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
                Log("Sẵn sàng. Chọn công cụ bên trái, thêm file rồi bấm Thực hiện.");
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
            if (tool.Input != FileKind.Pdf) lstTools.SelectedItem = tools[0];
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

        [DllImport("shell32.dll")] static extern int SHGetSpecialFolderLocation(IntPtr hwnd, int csidl, out IntPtr pidl);
        [DllImport("shell32.dll", EntryPoint = "SHGetFileInfoW")]
        static extern IntPtr SHGetFileInfoPidl(IntPtr pidl, uint attr, ref SHFILEINFO info, uint size, uint flags);

        // Biểu tượng của thư mục ảo (vd. CSIDL_DRIVES = 0x11 là "This PC")
        public static string KeySpecial(ImageList list, int csidl)
        {
            string key = "<csidl" + csidl + ">";
            if (list.Images.ContainsKey(key)) return key;
            IntPtr pidl;
            if (SHGetSpecialFolderLocation(IntPtr.Zero, csidl, out pidl) != 0) return null;
            var info = new SHFILEINFO();
            SHGetFileInfoPidl(pidl, 0, ref info, (uint)Marshal.SizeOf(info), SHGFI_ICON | SHGFI_SMALLICON | 0x8 /* SHGFI_PIDL */);
            Marshal.FreeCoTaskMem(pidl);
            if (info.hIcon == IntPtr.Zero) return null;
            using (var ic = Icon.FromHandle(info.hIcon)) list.Images.Add(key, ic.ToBitmap());
            DestroyIcon(info.hIcon);
            return key;
        }

        // Tên hiển thị giống Explorer (vd. "Local Disk (C:)", "This PC", "Downloads")
        public static string DisplayName(string path)
        {
            var info = new SHFILEINFO();
            if (SHGetFileInfo(path, 0, ref info, (uint)Marshal.SizeOf(info), 0x200) == IntPtr.Zero) return null;
            return string.IsNullOrEmpty(info.szDisplayName) ? null : info.szDisplayName;
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
                Func<int, int> S = Theme.S;
                f.Text = title;
                f.Font = Theme.Ui(10f);
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.StartPosition = FormStartPosition.CenterParent;
                f.MinimizeBox = f.MaximizeBox = false;
                f.ShowInTaskbar = false;
                f.AutoSize = true;
                f.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                f.Padding = new Padding(S(20), S(16), S(20), S(16));
                var flow = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false, Dock = DockStyle.Fill };
                flow.Controls.Add(new Label { Text = prompt, AutoSize = true, MaximumSize = new Size(S(380), 0), Margin = new Padding(0, 0, 0, S(10)) });
                var tb = new TextBox { Text = value, UseSystemPasswordChar = password };
                flow.Controls.Add(new FieldHost(tb) { Width = S(380), Margin = new Padding(0, 0, 0, S(18)) });
                var btns = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Size = new Size(S(380), S(36)), Margin = new Padding(0) };
                var cancel = new FButton("Hủy", null, ButtonKind.Standard) { DialogResult = DialogResult.Cancel, Width = S(100), Margin = new Padding(0) };
                var ok = new FButton("OK", null, ButtonKind.Primary) { DialogResult = DialogResult.OK, Width = S(100), Margin = new Padding(0, 0, S(8), 0) };
                btns.Controls.Add(cancel);
                btns.Controls.Add(ok);
                flow.Controls.Add(btns);
                f.Controls.Add(flow);
                f.AcceptButton = ok;
                f.CancelButton = cancel;
                Theme.Apply(f);
                f.Shown += (o, e) => { tb.Focus(); tb.SelectAll(); };
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
        readonly FileList list;
        readonly TextBox address;
        readonly ImageList icons;
        readonly SplitContainer split;
        readonly System.Windows.Forms.Timer refreshTimer;
        readonly Stack<string> history = new Stack<string>(), forward = new Stack<string>();
        readonly FButton btnBack, btnForward, btnUp;
        FileSystemWatcher watcher;
        string current;
        public Action<string[]> OpenFiles;

        static readonly string[] ColumnNames = { "Tên", "Ngày sửa", "Kích thước" };
        int sortCol;           // 0 tên, 1 ngày sửa, 2 kích thước
        bool sortDesc;
        int lastHeaderCol = -1;
        DateTime lastHeaderClick;

        void UpdateHeaders()
        {
            for (int i = 0; i < ColumnNames.Length; i++) list.Columns[i].Text = ColumnNames[i];
            list.Invalidate();
        }

        // Thư mục luôn đứng trước file; thư mục không có kích thước nên sắp theo tên
        List<T> Sorted<T>(IEnumerable<T> items, bool dirs) where T : FileSystemInfo
        {
            var byName = StringComparer.CurrentCultureIgnoreCase;
            if (sortCol == 1)
                return (sortDesc ? items.OrderByDescending(x => x.LastWriteTime) : items.OrderBy(x => x.LastWriteTime)).ToList();
            if (sortCol == 2 && !dirs)
                return (sortDesc ? items.OrderByDescending(x => ((FileInfo)(FileSystemInfo)x).Length) : items.OrderBy(x => ((FileInfo)(FileSystemInfo)x).Length)).ToList();
            if (sortCol == 2) return items.OrderBy(x => x.Name, byName).ToList();
            return (sortDesc ? items.OrderByDescending(x => x.Name, byName) : items.OrderBy(x => x.Name, byName)).ToList();
        }

        public string CurrentFolder { get { return current; } }

        public static bool Supported(string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            return ext == ".pdf" || ImageUtil.Extensions.Contains(ext);
        }

        public FileBrowser(float scale)
        {
            this.scale = scale;
            Font = Theme.Ui(10f);
            Theme.Bg(this, Role.Surface);
            icons = new ImageList { ColorDepth = ColorDepth.Depth32Bit, ImageSize = new Size(S(16), S(16)) };
            icons.Images.Add("<none>", new Bitmap(S(16), S(16)));

            // Thanh điều hướng: lùi, tiến, lên, làm mới, ô địa chỉ
            var bar = new TableLayoutPanel { Dock = DockStyle.Top, ColumnCount = 5, RowCount = 1, AutoSize = true, Padding = new Padding(S(8), S(8), S(8), S(6)) };
            Theme.Bg(bar, Role.Surface);
            for (int i = 0; i < 4; i++) bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            Func<string, string, EventHandler, FButton> btn = (glyph, hint, click) =>
            {
                var b = new FButton(null, glyph, ButtonKind.Subtle) { Margin = new Padding(0, 0, S(2), 0), Anchor = AnchorStyles.Left };
                b.Click += click;
                Theme.Tip(b, hint);
                return b;
            };
            btnBack = btn("", "Quay lại (Alt + ←)", (s, e) => GoBack());
            btnForward = btn("", "Tiến tới (Alt + →)", (s, e) => GoForward());
            btnUp = btn("", "Lên thư mục cha (Backspace)", (s, e) => GoUp());
            bar.Controls.Add(btnBack, 0, 0);
            bar.Controls.Add(btnForward, 1, 0);
            bar.Controls.Add(btnUp, 2, 0);
            bar.Controls.Add(btn("", "Làm mới (F5)", (s, e) => Navigate(current, false)), 3, 0);
            address = new TextBox();
            var addressHost = new FieldHost(address) { Dock = DockStyle.Fill, Margin = new Padding(S(6), 0, 0, 0), LeadingGlyph = "" };
            address.KeyDown += (s, e) =>
            {
                if (e.KeyCode != Keys.Enter) return;
                e.SuppressKeyPress = true;
                string p = address.Text.Trim().Trim('"');
                if (Directory.Exists(p)) Navigate(p, true);
                else if (File.Exists(p) && Supported(p) && OpenFiles != null) OpenFiles(new[] { p });
            };
            bar.Controls.Add(addressHost, 4, 0);

            split = new SplitContainer { Dock = DockStyle.Fill, FixedPanel = FixedPanel.Panel1, SplitterWidth = S(5) };
            Theme.Bg(split, Role.Surface);
            split.Paint += (s, e) =>
            {
                var r = split.SplitterRectangle;
                using (var pen = new Pen(Theme.P.Divider)) e.Graphics.DrawLine(pen, r.X + r.Width / 2, r.Y + S(4), r.X + r.Width / 2, r.Bottom - S(4));
            };
            tree = new NavTree
            {
                Dock = DockStyle.Fill, ImageList = icons, HideSelection = false, BorderStyle = BorderStyle.None,
                ItemHeight = S(32), Indent = S(14)
            };
            Theme.Bg(tree, Role.Surface);
            // Vẽ giống khung điều hướng của File Explorer (mũi tên, tô sáng khi rê chuột, vạch màu nhấn ở mục đang chọn)
            TreeStyle.Setup(tree, Separator);
            tree.BeforeSelect += (s, e) => { if (e.Node.Tag == Separator) e.Cancel = true; };
            tree.BeforeExpand += (s, e) => Populate(e.Node);
            tree.AfterSelect += (s, e) => { if (e.Action != TreeViewAction.Unknown && e.Node.Tag is string) Navigate((string)e.Node.Tag, true); };

            list = new FileList { Dock = DockStyle.Fill, Images = icons, FillColumn = 0 };
            Theme.Bg(list, Role.Surface);
            list.Columns.Add(new FileList.Column { Text = ColumnNames[0], Width = S(250) });
            list.Columns.Add(new FileList.Column { Text = ColumnNames[1], Width = S(140) });
            list.Columns.Add(new FileList.Column { Text = ColumnNames[2], Width = S(92), Right = true });
            int.TryParse(Settings.Get("sortcol", "0"), out sortCol);
            if (sortCol < 0 || sortCol >= ColumnNames.Length) sortCol = 0;
            sortDesc = Settings.Get("sortdesc", "0") == "1";
            list.SortState = col => col == sortCol ? (sortDesc ? 1 : 0) : -1;
            // Bấm tiêu đề cột: sắp xếp theo cột đó, bấm lại để đảo chiều; nhấn đúp luôn về A→Z
            list.HeaderClick += col =>
            {
                var now = DateTime.Now;
                bool dbl = col == lastHeaderCol && (now - lastHeaderClick).TotalMilliseconds <= SystemInformation.DoubleClickTime;
                lastHeaderCol = col;
                lastHeaderClick = dbl ? DateTime.MinValue : now;
                if (dbl) sortDesc = false;
                else if (col == sortCol) sortDesc = !sortDesc;
                else { sortCol = col; sortDesc = false; }
                Settings.Set("sortcol", sortCol.ToString());
                Settings.Set("sortdesc", sortDesc ? "1" : "0");
                UpdateHeaders();
                Navigate(current, false);
            };
            list.ItemDragStart += () =>
            {
                var paths = Selection();
                if (paths.Length > 0) list.DoDragDrop(new DataObject(DataFormats.FileDrop, paths), DragDropEffects.Copy);
            };
            list.ItemActivate += (s, e) => OpenSelected();
            list.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Back) GoUp();
                else if (e.KeyCode == Keys.F5) Navigate(current, false);
                else if (e.Alt && e.KeyCode == Keys.Left) GoBack();
                else if (e.Alt && e.KeyCode == Keys.Right) GoForward();
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
                try { split.SplitterDistance = S(210); } catch { }
            };
            AddRoots();
        }

        string[] Selection()
        {
            return list.SelectedRows.Select(r => r.Tag as string).Where(p => p != null).ToArray();
        }

        void OpenSelected()
        {
            var paths = Selection();
            if (paths.Length == 1 && Directory.Exists(paths[0])) { Navigate(paths[0], true); return; }
            var files = paths.Where(File.Exists).ToArray();
            if (files.Length > 0 && OpenFiles != null) OpenFiles(files);
        }

        static readonly object Separator = new object();
        const string ThisPc = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}";

        [DllImport("shell32.dll")]
        static extern int SHGetKnownFolderPath([MarshalAs(UnmanagedType.LPStruct)] Guid id, uint flags, IntPtr token, out IntPtr path);

        static string KnownFolder(string guid)
        {
            IntPtr p;
            if (SHGetKnownFolderPath(new Guid(guid), 0, IntPtr.Zero, out p) != 0) return null;
            string s = Marshal.PtrToStringUni(p);
            Marshal.FreeCoTaskMem(p);
            return s;
        }

        // Thư mục trong "Truy cập nhanh" (Quick access) của Explorer, theo đúng thứ tự
        static List<string> QuickAccessFolders()
        {
            var list = new List<string>();
            try
            {
                Func<object, string, BindingFlags, object[], object> call = (o, name, f, a) => o.GetType().InvokeMember(name, f, null, o, a);
                object shell = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));
                object ns = call(shell, "NameSpace", BindingFlags.InvokeMethod, new object[] { "shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}" });
                object items = call(ns, "Items", BindingFlags.InvokeMethod, null);
                int n = Convert.ToInt32(call(items, "Count", BindingFlags.GetProperty, null));
                for (int i = 0; i < n; i++)
                {
                    object it = call(items, "Item", BindingFlags.InvokeMethod, new object[] { i });
                    string path = call(it, "Path", BindingFlags.GetProperty, null) as string;
                    if (!string.IsNullOrEmpty(path) && Directory.Exists(path) &&
                        !list.Contains(path, StringComparer.OrdinalIgnoreCase)) list.Add(path);
                }
            }
            catch { }
            return list;
        }

        void AddRoots()
        {
            tree.BeginUpdate();
            // OneDrive (cá nhân và công ty), dùng biểu tượng đám mây của OneDrive như Explorer
            try
            {
                string cloudKey = null;
                using (var od = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\OneDrive"))
                {
                    string exe = od != null ? od.GetValue("OneDriveTrigger") as string : null;
                    if (!string.IsNullOrEmpty(exe) && File.Exists(exe)) cloudKey = ShellIcons.Key(icons, exe, false, true);
                }
                using (var accounts = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\OneDrive\Accounts"))
                    if (accounts != null)
                        foreach (string name in accounts.GetSubKeyNames().OrderBy(x => x == "Personal" ? 0 : 1))
                            using (var k = accounts.OpenSubKey(name))
                            {
                                string folder = k.GetValue("UserFolder") as string;
                                if (string.IsNullOrEmpty(folder)) continue;
                                string org = k.GetValue("DisplayName") as string;
                                AddNode(tree.Nodes, name == "Personal" ? "OneDrive - Personal" : "OneDrive - " + (string.IsNullOrEmpty(org) ? Path.GetFileName(folder) : org), folder, true);
                                if (cloudKey != null && tree.Nodes.Count > 0 && tree.Nodes[tree.Nodes.Count - 1].Tag as string == folder)
                                    tree.Nodes[tree.Nodes.Count - 1].ImageKey = tree.Nodes[tree.Nodes.Count - 1].SelectedImageKey = cloudKey;
                            }
            }
            catch { }
            AddSeparator();

            // Các thư mục quen thuộc + thư mục ghim trong Truy cập nhanh
            var known = new[]
            {
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                KnownFolder("374DE290-123F-4565-9164-39C4925E467B"),   // Downloads
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                Environment.GetFolderPath(Environment.SpecialFolder.MyPictures),
                Environment.GetFolderPath(Environment.SpecialFolder.MyMusic),
                Environment.GetFolderPath(Environment.SpecialFolder.MyVideos)
            }.Where(p => !string.IsNullOrEmpty(p)).ToList();
            foreach (string p in known.Concat(QuickAccessFolders().Where(q => !known.Contains(q, StringComparer.OrdinalIgnoreCase))))
                AddNode(tree.Nodes, null, p, false);
            AddSeparator();

            // This PC và các ổ đĩa
            string pcKey = ShellIcons.KeySpecial(icons, 0x11);
            var pc = new TreeNode(ShellIcons.DisplayName(ThisPc) ?? "This PC") { ImageKey = pcKey, SelectedImageKey = pcKey };
            tree.Nodes.Add(pc);
            foreach (var d in DriveInfo.GetDrives())
            {
                try { if (d.IsReady) AddNode(pc.Nodes, null, d.RootDirectory.FullName, true); } catch { }
            }
            pc.Expand();
            tree.EndUpdate();
        }

        void AddNode(TreeNodeCollection parent, string text, string path, bool expandable)
        {
            if (string.IsNullOrEmpty(path) || !Directory.Exists(path)) return;
            string key = ShellIcons.Key(icons, path, true, true);
            var n = new TreeNode(text ?? ShellIcons.DisplayName(path) ?? Path.GetFileName(path.TrimEnd('\\'))) { Tag = path, ImageKey = key, SelectedImageKey = key };
            if (expandable) n.Nodes.Add(new TreeNode("..."));
            parent.Add(n);
        }

        // Tô sáng mục trên cây ứng với thư mục đang xem (nếu là mục gốc hoặc ổ đĩa)
        void SelectInTree(string path)
        {
            TreeNode hit = null;
            foreach (TreeNode n in tree.Nodes)
            {
                if (string.Equals(n.Tag as string, path, StringComparison.OrdinalIgnoreCase)) { hit = n; break; }
                foreach (TreeNode c in n.Nodes)
                    if (string.Equals(c.Tag as string, path, StringComparison.OrdinalIgnoreCase)) { hit = c; break; }
                if (hit != null) break;
            }
            tree.SelectedNode = hit;
        }

        void AddSeparator()
        {
            tree.Nodes.Add(new TreeNode("") { Tag = Separator, ImageKey = "<none>", SelectedImageKey = "<none>" });
        }

        void Populate(TreeNode node)
        {
            if (!(node.Tag is string) || node.Nodes.Count != 1 || node.Nodes[0].Tag != null) return;
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
            if (history.Count == 0) return;
            forward.Push(current);
            Navigate(history.Pop(), false);
        }

        void GoForward()
        {
            if (forward.Count == 0) return;
            history.Push(current);
            Navigate(forward.Pop(), false);
        }

        public void Navigate(string path, bool addHistory)
        {
            if (string.IsNullOrEmpty(path) || !Directory.Exists(path)) return;
            bool refresh = string.Equals(current, path, StringComparison.OrdinalIgnoreCase);
            if (addHistory && current != null && !refresh) { history.Push(current); forward.Clear(); }
            var keep = refresh ? new HashSet<string>(Selection(), StringComparer.OrdinalIgnoreCase) : new HashSet<string>();
            int top = refresh ? list.TopIndex : 0;
            current = path;
            address.Text = path;
            list.BeginUpdate();
            list.Clear();
            try
            {
                var di = new DirectoryInfo(path);
                string dirKey = ShellIcons.Key(icons, "", true, false);
                foreach (var d in Sorted(SubDirs(path), true))
                    list.Add(new[] { d.Name, d.LastWriteTime.ToString("dd/MM/yyyy HH:mm"), "" }, dirKey, d.FullName, keep.Contains(d.FullName));
                foreach (var f in Sorted(di.GetFiles().Where(f => IsShown(f) && Supported(f.Name)), false))
                    list.Add(new[] { f.Name, f.LastWriteTime.ToString("dd/MM/yyyy HH:mm"), Job.SizeText(f.Length) },
                             ShellIcons.Key(icons, f.FullName, false, false), f.FullName, keep.Contains(f.FullName));
            }
            catch (Exception ex) { list.Add(new[] { "(" + ex.Message + ")", "", "" }, null, null, false); }
            list.EndUpdate();
            list.TopIndex = top;
            SelectInTree(path);
            btnBack.Enabled = history.Count > 0;
            btnForward.Enabled = forward.Count > 0;
            btnUp.Enabled = Directory.GetParent(path) != null;
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
    class Workspace : Panel, IThemed
    {
        public readonly List<PdfDoc> Docs = new List<PdfDoc>();
        public readonly HashSet<PageRef> Selected = new HashSet<PageRef>();
        public PdfDoc Active;
        public Session Session;
        public event Action Changed;
        public Action<PdfDoc> SaveRequested, CloseRequested;
        public Action<string[], PdfDoc, int> FilesDropped;     // doc null = tạo tài liệu mới
        public Action<PageRef> PreviewRequested;
        public Action OpenRequested;
        public ContextMenuStrip PageMenu, DocMenu;
        int newCounter;

        readonly float scale;
        int box;
        int S(int v) { return (int)(v * scale); }
        readonly FButton openBtn;

        class DocLayout
        {
            public PdfDoc Doc;
            public Rectangle Block, Header, SaveBtn, CloseBtn;
            public readonly List<Rectangle> Cells = new List<Rectangle>();
        }
        readonly List<DocLayout> layout = new List<DocLayout>();

        Point downPt;
        bool dragArmed;
        PageRef pendingSingle, anchor, hoverPage;
        PdfDoc hoverDoc;
        int hoverBtn;       // 1 = Lưu, 2 = Đóng
        bool dropActive;
        PdfDoc dropDoc;
        int dropIndex;
        const string PagesFormat = "PdfToolPages";
        const TextFormatFlags Keep = TextFormatFlags.PreserveGraphicsTranslateTransform | TextFormatFlags.PreserveGraphicsClipping;

        public Workspace(float scale)
        {
            this.scale = scale;
            box = S(150);
            SetStyle(ControlStyles.Selectable | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
            AutoScroll = true;
            AllowDrop = true;
            TabStop = true;
            BackColor = Theme.P.Canvas;
            openBtn = new FButton("Mở tài liệu...", "", ButtonKind.Primary);
            openBtn.Size = new Size(Math.Max(S(170), openBtn.GetPreferredSize(Size.Empty).Width), S(38));
            openBtn.Click += (s, e) => { if (OpenRequested != null) OpenRequested(); };
            Controls.Add(openBtn);
        }

        public void ApplyTheme()
        {
            BackColor = Theme.P.Canvas;
            Theme.Native(this);
            Invalidate();
        }

        void OnChanged()
        {
            var h = Changed;
            if (h != null) h();
        }

        int CellW { get { return box + S(28); } }
        int CellH { get { return box + S(36); } }

        public void Relayout()
        {
            if (Active != null && !Docs.Contains(Active)) Active = Docs.FirstOrDefault();
            Selected.RemoveWhere(p => !Docs.Any(d => d.Pages.Contains(p)));
            layout.Clear();
            int m = S(18), y = m;
            int width = Math.Max(CellW + 2 * m + S(24), ClientSize.Width);
            int cols = Math.Max(1, (width - 2 * m - S(24)) / CellW);
            int hh = S(52);
            foreach (var d in Docs)
            {
                var L = new DocLayout { Doc = d };
                int rows = Math.Max(1, (d.Pages.Count + cols - 1) / cols);
                L.Block = new Rectangle(m, y, width - 2 * m, hh + S(12) + rows * CellH + S(8));
                L.Header = new Rectangle(m, y, width - 2 * m, hh);
                int bh = S(32);
                L.CloseBtn = new Rectangle(L.Header.Right - S(12) - bh, y + (hh - bh) / 2, bh, bh);
                L.SaveBtn = new Rectangle(L.CloseBtn.Left - S(6) - S(80), y + (hh - bh) / 2, S(80), bh);
                for (int i = 0; i < d.Pages.Count; i++)
                    L.Cells.Add(new Rectangle(m + S(12) + (i % cols) * CellW, y + hh + S(12) + (i / cols) * CellH, CellW, CellH));
                layout.Add(L);
                y = L.Block.Bottom + m;
            }
            AutoScrollMinSize = new Size(0, Docs.Count > 0 ? y : 0);
            openBtn.Visible = Docs.Count == 0;
            openBtn.Location = new Point((ClientSize.Width - openBtn.Width) / 2, ClientSize.Height / 2 + S(56));
            Invalidate();
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            Relayout();
        }

        // ---------------- vẽ ----------------
        static void DrawText(Graphics g, string s, Font f, Rectangle r, Color c, TextFormatFlags flags)
        {
            TextRenderer.DrawText(g, s, f, r, c, flags | Keep);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var p = Theme.P;
            var g = e.Graphics;
            g.Clear(p.Canvas);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;

            if (Docs.Count == 0) { DrawEmpty(g); if (dropActive) DrawNewDocTarget(g); return; }

            var state = g.Save();
            g.TranslateTransform(AutoScrollPosition.X, AutoScrollPosition.Y);
            var clip = e.ClipRectangle;
            clip.Offset(-AutoScrollPosition.X, -AutoScrollPosition.Y);
            foreach (var L in layout)
                if (L.Block.IntersectsWith(clip)) DrawDoc(g, L, clip);

            // Vị trí thả
            if (dropActive && dropDoc != null)
            {
                var L = layout.FirstOrDefault(x => x.Doc == dropDoc);
                if (L != null && L.Cells.Count > 0)
                {
                    Rectangle c = dropIndex < L.Cells.Count ? L.Cells[dropIndex] : L.Cells[L.Cells.Count - 1];
                    int x = dropIndex < L.Cells.Count ? c.Left : c.Right;
                    int y0 = c.Top + S(4), y1 = c.Top + S(10) + box;
                    using (var pen = new Pen(p.Accent, S(3)) { StartCap = LineCap.Round, EndCap = LineCap.Round }) g.DrawLine(pen, x, y0, x, y1);
                    using (var b = new SolidBrush(p.Accent))
                    {
                        g.FillEllipse(b, x - S(5), y0 - S(5), S(10), S(10));
                        g.FillEllipse(b, x - S(5), y1 - S(5), S(10), S(10));
                    }
                }
            }
            g.Restore(state);
            if (dropActive && dropDoc == null) DrawNewDocTarget(g);
        }

        void DrawDoc(Graphics g, DocLayout L, Rectangle clip)
        {
            var p = Theme.P;
            var d = L.Doc;
            bool active = d == Active;
            Theme.FillRound(g, new RectangleF(L.Block.X + 0.5f, L.Block.Y + 0.5f, L.Block.Width - 1, L.Block.Height - 1), S(10),
                            p.Surface, active ? p.Accent : p.Border, active ? S(2) : 1);

            var h = L.Header;
            var ic = new Rectangle(h.X + S(16), h.Y + (h.Height - S(30)) / 2, S(30), S(30));
            Theme.Glyph(g, "", ic, p.AccentText, 16f, false);
            int x = ic.Right + S(10);
            var tf = Theme.UiSemi(11f);
            int maxTitle = Math.Max(S(60), L.SaveBtn.Left - x - S(200));
            int tw = Math.Min(TextRenderer.MeasureText(g, d.Name, tf, Size.Empty, TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix).Width + S(2), maxTitle);
            DrawText(g, d.Name, tf, new Rectangle(x, h.Y, tw, h.Height), p.Text, Theme.TextLeft);
            x += tw + S(12);
            x = Chip(g, x, h, d.Pages.Count + " trang", p.TextMuted, Color.FromArgb(p.Dark ? 34 : 14, p.Text)) + S(6);
            if (d.Modified) Chip(g, x, h, "Chưa lưu", p.Warning, p.WarningBg);

            HeaderButton(g, L.SaveBtn, "", "Lưu", d.Modified, hoverDoc == d && hoverBtn == 1);
            HeaderButton(g, L.CloseBtn, "", null, false, hoverDoc == d && hoverBtn == 2);
            using (var pen = new Pen(p.Divider)) g.DrawLine(pen, L.Block.X + 1, h.Bottom, L.Block.Right - 2, h.Bottom);

            for (int i = 0; i < d.Pages.Count; i++)
                if (L.Cells[i].IntersectsWith(clip)) DrawPage(g, d.Pages[i], L.Cells[i], i + 1);
        }

        int Chip(Graphics g, int x, Rectangle h, string text, Color fg, Color bg)
        {
            var f = Theme.Ui(8.5f);
            int tw = TextRenderer.MeasureText(g, text, f, Size.Empty, TextFormatFlags.NoPadding).Width;
            var r = new Rectangle(x, h.Y + (h.Height - S(22)) / 2, tw + S(18), S(22));
            Theme.FillRound(g, r, S(11), bg, Color.Empty, 0);
            DrawText(g, text, f, r, fg, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix);
            return r.Right;
        }

        void HeaderButton(Graphics g, Rectangle r, string glyph, string text, bool primary, bool hot)
        {
            var p = Theme.P;
            Color fill = primary ? (hot ? p.AccentHover : p.Accent) : (hot ? (text != null ? p.ButtonHover : p.Hover) : (text != null ? p.ButtonBg : Color.Empty));
            Color border = !primary && text != null ? p.ButtonBorder : Color.Empty;
            Color fg = primary ? p.OnAccent : p.Text;
            Theme.FillRound(g, new RectangleF(r.X + 0.5f, r.Y + 0.5f, r.Width - 1, r.Height - 1), S(4), fill, border, 1);
            if (text == null) { Theme.Glyph(g, glyph, r, fg, 10f, false); return; }
            var f = Theme.Ui(9.5f);
            int tw = TextRenderer.MeasureText(g, text, f, Size.Empty, TextFormatFlags.NoPadding).Width;
            int x = r.X + (r.Width - (S(16) + S(6) + tw)) / 2;
            Theme.Glyph(g, glyph, new RectangleF(x, r.Y, S(16), r.Height), fg, 10f, false);
            DrawText(g, text, f, new Rectangle(x + S(22), r.Y, tw + S(2), r.Height), fg, Theme.TextLeft);
        }

        void DrawPage(Graphics g, PageRef pg, Rectangle cell, int number)
        {
            var p = Theme.P;
            bool sel = Selected.Contains(pg), hot = pg == hoverPage;
            var r0 = Rectangle.Inflate(cell, -S(4), -S(2));
            if (sel) Theme.FillRound(g, new RectangleF(r0.X + 1, r0.Y + 1, r0.Width - 2, r0.Height - 2), S(8), p.Selection, p.Accent, S(2));
            else if (hot) Theme.FillRound(g, r0, S(8), p.Hover, Color.Empty, 0);

            Bitmap bmp = Session.Thumb(pg);
            double w = bmp != null ? bmp.Width : 210, h = bmp != null ? bmp.Height : 297;
            bool swap = pg.Rot % 180 != 0;
            double vw = swap ? h : w, vh = swap ? w : h;
            double s = Math.Min(box / vw, box / vh);
            int dw = Math.Max(4, (int)(vw * s)), dh = Math.Max(4, (int)(vh * s));
            var r = new Rectangle(cell.X + (cell.Width - dw) / 2, cell.Y + S(10) + (box - dh) / 2, dw, dh);
            // Bóng đổ mềm dưới trang
            var mode = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.None;
            for (int k = 3; k >= 1; k--)
                using (var b = new SolidBrush(Color.FromArgb(p.Dark ? 70 / k : 30 / k, 0, 0, 0)))
                    g.FillRectangle(b, r.X - k + 1, r.Y - k + S(2), r.Width + 2 * k - 2, r.Height + 2 * k - 2);
            g.FillRectangle(Brushes.White, r);
            if (bmp != null)
            {
                var st = g.Save();
                g.TranslateTransform(r.X + r.Width / 2f, r.Y + r.Height / 2f);
                g.RotateTransform(pg.Rot);
                float uw = swap ? r.Height : r.Width, uh = swap ? r.Width : r.Height;
                g.DrawImage(bmp, -uw / 2, -uh / 2, uw, uh);
                g.Restore(st);
            }
            else Theme.Glyph(g, "", r, Color.FromArgb(255, 0xC8, 0xC8, 0xC8), 18f, false);
            if (!p.Dark) using (var pen = new Pen(Color.FromArgb(30, 0, 0, 0))) g.DrawRectangle(pen, r.X, r.Y, r.Width - 1, r.Height - 1);
            g.SmoothingMode = mode;
            DrawText(g, number.ToString(), sel ? Theme.UiSemi(9f) : Theme.Ui(9f), new Rectangle(cell.X, cell.Y + S(12) + box, cell.Width, S(20)),
                 sel ? p.AccentText : p.TextMuted, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
        }

        void DrawEmpty(Graphics g)
        {
            var p = Theme.P;
            var r = Rectangle.Inflate(ClientRectangle, -S(28), -S(28));
            if (r.Width < 50 || r.Height < 50) return;
            using (var pen = new Pen(p.Dark ? Color.FromArgb(255, 0x44, 0x44, 0x44) : Color.FromArgb(255, 0xC6, 0xC9, 0xCE), 1.5f) { DashStyle = DashStyle.Dash })
            using (var path = Theme.Round(r, S(14))) g.DrawPath(pen, path);
            int cx = ClientSize.Width / 2, cy = ClientSize.Height / 2;
            var circle = new RectangleF(cx - S(44), cy - S(128), S(88), S(88));
            using (var b = new SolidBrush(Color.FromArgb(p.Dark ? 40 : 24, p.Accent))) g.FillEllipse(b, circle);
            Theme.Glyph(g, "", circle, p.AccentText, 30f, false);
            TextRenderer.DrawText(g, "Kéo thả tài liệu PDF hoặc ảnh vào đây", Theme.UiSemi(15f), new Rectangle(0, cy - S(28), ClientSize.Width, S(34)), p.Text,
                                  TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
            TextRenderer.DrawText(g, "từ danh sách bên trái hoặc từ File Explorer. Hỗ trợ PDF, JPG, PNG, BMP, GIF, TIFF.", Theme.Ui(10f),
                                  new Rectangle(0, cy + S(8), ClientSize.Width, S(26)), p.TextMuted,
                                  TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
        }

        void DrawNewDocTarget(Graphics g)
        {
            var p = Theme.P;
            using (var b = new SolidBrush(Color.FromArgb(p.Dark ? 26 : 16, p.Accent))) g.FillRectangle(b, ClientRectangle);
            var r = Rectangle.Inflate(ClientRectangle, -S(10), -S(10));
            using (var pen = new Pen(p.Accent, S(2)) { DashStyle = DashStyle.Dash })
            using (var path = Theme.Round(r, S(12))) g.DrawPath(pen, path);
            var f = Theme.UiSemi(10f);
            string text = "Thả để tạo tài liệu mới";
            int tw = TextRenderer.MeasureText(g, text, f, Size.Empty, TextFormatFlags.NoPadding).Width;
            var lbl = new Rectangle(r.X + r.Width / 2 - (tw + S(56)) / 2, r.Bottom - S(58), tw + S(56), S(38));
            Theme.FillRound(g, lbl, S(19), p.Accent, Color.Empty, 0);
            Theme.Glyph(g, "", new RectangleF(lbl.X + S(16), lbl.Y, S(16), lbl.Height), p.OnAccent, 10f, false);
            TextRenderer.DrawText(g, text, f, new Rectangle(lbl.X + S(40), lbl.Y, tw + S(4), lbl.Height), p.OnAccent, Theme.TextLeft);
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
            if (e.Button == MouseButtons.None) { UpdateHover(e.Location); return; }
            if (!dragArmed || e.Button != MouseButtons.Left) return;
            var ds = SystemInformation.DragSize;
            if (Math.Abs(e.X - downPt.X) < ds.Width && Math.Abs(e.Y - downPt.Y) < ds.Height) return;
            dragArmed = false;
            pendingSingle = null;
            DoDragDrop(new DataObject(PagesFormat, "pages"), DragDropEffects.Move | DragDropEffects.Copy);
            dropActive = false;
            Invalidate();
        }

        // Tô sáng trang và nút dưới con trỏ chuột
        void UpdateHover(Point client)
        {
            Point p = ToContent(client);
            var L = DocAt(p);
            PageRef page = null;
            int btn = 0;
            if (L != null)
            {
                int i = PageAt(L, p);
                if (i >= 0) page = L.Doc.Pages[i];
                btn = L.SaveBtn.Contains(p) ? 1 : L.CloseBtn.Contains(p) ? 2 : 0;
            }
            PdfDoc doc = L != null ? L.Doc : null;
            if (page == hoverPage && btn == hoverBtn && (btn == 0 || doc == hoverDoc)) return;
            hoverPage = page;
            hoverBtn = btn;
            hoverDoc = doc;
            Cursor = btn != 0 ? Cursors.Hand : Cursors.Default;
            Invalidate();
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            if (hoverPage == null && hoverBtn == 0) return;
            hoverPage = null;
            hoverBtn = 0;
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
        readonly FButton prev, next;

        public PreviewForm(Session session, List<PageRef> pages, int index, string title)
        {
            this.session = session;
            this.pages = pages;
            this.index = index;
            Func<int, int> S = Theme.S;
            Text = "Xem trước - " + title;
            Font = Theme.Ui(10f);
            var wa = Screen.FromPoint(Cursor.Position).WorkingArea;
            Size = new Size(Math.Min(wa.Width, (int)(wa.Height * 0.8)), (int)(wa.Height * 0.9));
            StartPosition = FormStartPosition.CenterParent;
            KeyPreview = true;
            ShowInTaskbar = false;
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            pic = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, Padding = new Padding(S(16)) };
            Theme.Bg(pic, Role.Canvas);
            var bar = new FlowLayoutPanel { Dock = DockStyle.Top, Height = S(54), WrapContents = false, Padding = new Padding(S(12), S(10), S(12), S(8)) };
            Theme.Bg(bar, Role.Surface);
            prev = new FButton("Trang trước", "\uE72B", ButtonKind.Standard);
            next = new FButton("Trang sau", "\uE72A", ButtonKind.Standard);
            prev.Margin = next.Margin = new Padding(0, 0, S(8), 0);
            lbl = new Label { AutoSize = true, Font = Theme.UiSemi(10.5f), Margin = new Padding(S(10), S(8), 0, 0) };
            prev.Click += (o, e) => Go(-1);
            next.Click += (o, e) => Go(1);
            Theme.Tip(prev, "Trang trước (← hoặc Page Up)");
            Theme.Tip(next, "Trang sau (→ hoặc Page Down)");
            bar.Controls.AddRange(new Control[] { prev, next, lbl });
            var line = new Panel { Dock = DockStyle.Top, Height = 1 };
            Theme.Bg(line, Role.Divider);
            Controls.Add(pic);
            Controls.Add(line);
            Controls.Add(bar);
            KeyDown += (o, e) =>
            {
                if (e.KeyCode == Keys.Left || e.KeyCode == Keys.PageUp) Go(-1);
                else if (e.KeyCode == Keys.Right || e.KeyCode == Keys.PageDown) Go(1);
                else if (e.KeyCode == Keys.Escape) Close();
            };
            Theme.Apply(this);
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
            prev.Enabled = index > 0;
            next.Enabled = index < pages.Count - 1;
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
    // Ô "Ủng hộ cốc cafe" ở góc dưới bên trái (mã QR chuyển khoản)
    // ------------------------------------------------------------------
    class DonatePanel : Panel, IThemed
    {
        static Image card;   // ảnh thẻ VietQR nhúng trong exe (donate.png)
        readonly Bitmap qr;
        bool hover;

        static Image Card
        {
            get
            {
                if (card == null)
                    using (var s = typeof(DonatePanel).Assembly.GetManifestResourceStream("PdfTool.donate.png"))
                        if (s != null) card = new Bitmap(Image.FromStream(s));
                return card;
            }
        }

        static int S(int v) { return Theme.S(v); }

        public DonatePanel(float scale)
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
            BackColor = Color.Transparent;
            Dock = DockStyle.Bottom;
            Height = S(86);
            Cursor = Cursors.Hand;
            if (Card == null) { Visible = false; return; }

            // Cắt phần mã QR từ ảnh thẻ để hiển thị nhỏ
            qr = new Bitmap(S(52), S(52));
            using (var g = Graphics.FromImage(qr))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.Clear(Color.White);
                g.DrawImage(Card, new Rectangle(0, 0, qr.Width, qr.Height), new Rectangle(190, 220, 595, 595), GraphicsUnit.Pixel);
            }
            Theme.Tip(this, "Nhấn để phóng to mã QR, quét bằng ứng dụng ngân hàng để ủng hộ tác giả");
        }

        public void ApplyTheme() { Invalidate(); }

        protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnClick(EventArgs e) { base.OnClick(e); ShowCard(); }

        protected override void OnPaint(PaintEventArgs e)
        {
            if (qr == null) return;
            var g = e.Graphics;
            var p = Theme.P;
            var r = new RectangleF(S(4) + 0.5f, S(8) + 0.5f, Width - S(8) - 1, Height - S(14) - 1);
            Color bg = hover ? (p.Dark ? Color.FromArgb(255, 0x3E, 0x32, 0x22) : Color.FromArgb(255, 0xFF, 0xF1, 0xD6)) : p.DonateBg;
            Theme.FillRound(g, r, S(8), bg, p.DonateBorder, 1);
            var q = new RectangleF(r.X + S(8), r.Y + (r.Height - S(56)) / 2, S(56), S(56));
            Theme.FillRound(g, q, S(4), Color.White, Color.Empty, 0);
            g.InterpolationMode = InterpolationMode.NearestNeighbor;
            g.DrawImage(qr, q.X + S(2), q.Y + S(2), S(52), S(52));
            int x = (int)q.Right + S(10), y = (int)r.Y + S(5), w = (int)r.Right - x - S(4);
            TextRenderer.DrawText(g, "☕ Ủng hộ cốc cafe", Theme.UiSemi(9.5f), new Rectangle(x, y, w, S(19)), p.DonateTitle, Theme.TextLeft);
            TextRenderer.DrawText(g, "BIDV · 4260301070", Theme.Ui(8.5f), new Rectangle(x, y + S(19), w, S(16)), p.Text, Theme.TextLeft);
            TextRenderer.DrawText(g, "NGO QUANG HUY", Theme.Ui(8.5f), new Rectangle(x, y + S(34), w, S(16)), p.TextMuted, Theme.TextLeft);
            TextRenderer.DrawText(g, "Nhấn để phóng to mã QR", Theme.Ui(8f), new Rectangle(x, y + S(49), w, S(15)), p.TextMuted, Theme.TextLeft);
        }

        void ShowCard()
        {
            using (var f = new Form
            {
                Text = "Ủng hộ cốc cafe", FormBorderStyle = FormBorderStyle.FixedDialog, MinimizeBox = false, MaximizeBox = false,
                StartPosition = FormStartPosition.CenterParent, ShowInTaskbar = false,
                ClientSize = new Size(S(400), S(497)), Padding = new Padding(S(12))
            })
            {
                var pic = new PictureBox { Image = Card, Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom };
                Theme.Bg(pic, Role.Window);
                f.Controls.Add(pic);
                f.KeyPreview = true;
                f.KeyDown += (o, e) => { if (e.KeyCode == Keys.Escape) f.Close(); };
                Theme.Apply(f);
                f.ShowDialog(FindForm());
            }
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
        readonly FlowLayoutPanel bar;
        readonly Panel statusBar;
        readonly Label status;
        readonly List<FButton> needSelection = new List<FButton>(), needDoc = new List<FButton>();
        FButton bMerge, bCloseAll;
        public Action<string> SendToToolbox;

        public CreatorView(float scale)
        {
            this.scale = scale;
            Font = Theme.Ui(10f);
            Theme.Bg(this, Role.Window);

            split = new SplitContainer { Dock = DockStyle.Fill, FixedPanel = FixedPanel.Panel1, SplitterWidth = S(5) };
            Theme.Bg(split, Role.Surface);
            Theme.Bg(split.Panel1, Role.Surface);
            Theme.Bg(split.Panel2, Role.Canvas);
            split.Paint += (o, e) =>
            {
                var r = split.SplitterRectangle;
                using (var pen = new Pen(Theme.P.Border)) e.Graphics.DrawLine(pen, r.Right - 1, r.Y, r.Right - 1, r.Bottom);
            };
            browser = new FileBrowser(scale) { Dock = DockStyle.Fill };
            browser.OpenFiles = files => OpenFiles(files, null, -1);
            var donate = new DonatePanel(scale) { Padding = new Padding(0) };
            var donateHost = new Panel { Dock = DockStyle.Bottom, Height = donate.Height, Padding = new Padding(S(4), 0, S(4), 0) };
            Theme.Bg(donateHost, Role.Surface);
            donate.Dock = DockStyle.Fill;
            donateHost.Controls.Add(donate);
            split.Panel1.Controls.Add(browser);
            split.Panel1.Controls.Add(donateHost);

            ws = new Workspace(scale) { Dock = DockStyle.Fill, Session = session };
            ws.SaveRequested = d => SaveDoc(d);
            ws.CloseRequested = d => CloseDoc(d);
            ws.FilesDropped = (f, d, i) => OpenFiles(f, d, i);
            ws.PreviewRequested = Preview;
            ws.OpenRequested = OpenDialog;
            ws.Changed += UpdateUi;
            session.ThumbsChanged = () =>
            {
                try { if (ws.IsHandleCreated) ws.BeginInvoke(new Action(ws.Invalidate)); } catch { }
            };

            // Thanh công cụ
            bar = new FlowLayoutPanel { Dock = DockStyle.Top, Height = S(54), WrapContents = false, Padding = new Padding(S(10), S(10), S(10), S(8)) };
            Theme.Bg(bar, Role.Surface);
            bar.Paint += (o, e) => { using (var pen = new Pen(Theme.P.Divider)) e.Graphics.DrawLine(pen, 0, bar.Height - 1, bar.Width, bar.Height - 1); };
            Func<string, string, string, Action, FButton> add = (text, glyph, tip, act) =>
            {
                var b = new FButton(text, glyph, ButtonKind.Subtle) { Compactable = text != null, Margin = new Padding(0, 0, S(2), 0) };
                b.Click += (o, e) => act();
                Theme.Tip(b, tip);
                bar.Controls.Add(b);
                return b;
            };
            add("Mở", "\uE838", "Mở file PDF hoặc ảnh (Ctrl+O)", OpenDialog);
            needDoc.Add(add("Lưu", "\uE74E", "Lưu tài liệu đang chọn (Ctrl+S)", () => SaveDoc(ws.Active)));
            bar.Controls.Add(new Separator());
            bMerge = add("Gộp tất cả", "\uEA3C", "Gộp tất cả tài liệu thành một", ws.MergeAll);
            needSelection.Add(add("Tách trang", "\uE8C6", "Chuyển các trang đang chọn thành tài liệu mới", ws.ExtractSelected));
            bar.Controls.Add(new Separator());
            var rotL = add(null, "\uE7AD", "Xoay trái 90° các trang đang chọn", () => ws.RotateSelected(270));
            rotL.MirrorGlyph = true;
            needSelection.Add(rotL);
            needSelection.Add(add(null, "\uE7AD", "Xoay phải 90° các trang đang chọn", () => ws.RotateSelected(90)));
            needSelection.Add(add("Xóa", "\uE74D", "Xóa các trang đang chọn (Delete)", ws.DeleteSelected));
            needSelection.Add(add("Xem", "\uE890", "Xem trang đang chọn cỡ lớn (Enter hoặc nhấn đúp)", PreviewSelected));
            bar.Controls.Add(new Separator());
            add(null, "\uE71F", "Thu nhỏ ảnh trang (Ctrl + lăn chuột)", () => ws.Zoom(-S(30)));
            add(null, "\uE8A3", "Phóng to ảnh trang (Ctrl + lăn chuột)", () => ws.Zoom(S(30)));
            bar.Controls.Add(new Separator());
            needDoc.Add(add("Nén, OCR, mật khẩu...", "\uE90F", "Mở tài liệu đang chọn trong Hộp công cụ (nén, OCR, đánh số trang, mật khẩu...)", SendActiveToToolbox));
            bCloseAll = add("Đóng tất cả", "\uE711", "Đóng tất cả tài liệu", CloseAll);
            bar.Resize += (o, e) => FitToolbar();

            ws.PageMenu = Theme.Menu();
            Theme.Item(ws.PageMenu, "Xem trước", "\uE890", (o, e) => PreviewSelected());
            Theme.Item(ws.PageMenu, "Xoay trái", "\uE7A7", (o, e) => ws.RotateSelected(270));
            Theme.Item(ws.PageMenu, "Xoay phải", "\uE7AD", (o, e) => ws.RotateSelected(90));
            Theme.Item(ws.PageMenu, "Tách thành tài liệu mới", "\uE8C6", (o, e) => ws.ExtractSelected());
            ws.PageMenu.Items.Add(new ToolStripSeparator());
            Theme.Item(ws.PageMenu, "Xóa", "\uE74D", (o, e) => ws.DeleteSelected());
            ws.DocMenu = Theme.Menu();
            Theme.Item(ws.DocMenu, "Lưu...", "\uE74E", (o, e) => SaveDoc(ws.Active));
            Theme.Item(ws.DocMenu, "Đổi tên...", "\uE8AC", (o, e) => RenameActive());
            Theme.Item(ws.DocMenu, "Chọn tất cả trang", "\uE8B3", (o, e) => ws.SelectAll());
            Theme.Item(ws.DocMenu, "Mở trong Hộp công cụ...", "\uE90F", (o, e) => SendActiveToToolbox());
            ws.DocMenu.Items.Add(new ToolStripSeparator());
            Theme.Item(ws.DocMenu, "Đóng", "\uE711", (o, e) => CloseDoc(ws.Active));

            // Thanh trạng thái
            statusBar = new Panel { Dock = DockStyle.Bottom, Height = S(30) };
            Theme.Bg(statusBar, Role.Surface);
            statusBar.Paint += (o, e) => { using (var pen = new Pen(Theme.P.Divider)) e.Graphics.DrawLine(pen, 0, 0, statusBar.Width, 0); };
            status = new Label { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(S(12), 0, 0, 0), AutoEllipsis = true, Font = Theme.Ui(9f) };
            Theme.Fg(status, TextRole.Muted);
            statusBar.Controls.Add(status);

            split.Panel2.Controls.Add(ws);
            split.Panel2.Controls.Add(bar);
            Controls.Add(split);
            Controls.Add(statusBar);

            Load += (o, e) =>
            {
                try
                {
                    // Vùng soạn luôn rộng tối thiểu ~560px; khung duyệt file không hẹp quá 320px
                    split.Panel1MinSize = S(320);
                    split.Panel2MinSize = S(560);
                    split.SplitterDistance = Math.Min(S(int.Parse(Settings.Get("browserwidth", "640"))), split.Width - S(560) - split.SplitterWidth);
                }
                catch { }
                browser.Navigate(Settings.Get("browser", Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)), false);
                UpdateUi();
                FitToolbar();
            };
        }

        public bool HasUnsaved { get { return ws.Docs.Any(d => d.Modified); } }

        public void SaveSettings()
        {
            Settings.Set("browserwidth", ((int)(split.SplitterDistance / scale)).ToString());
        }

        public void Cleanup() { session.Dispose(); }

        // Thanh công cụ hẹp: ẩn chữ, chỉ còn biểu tượng (vẫn có chú thích khi rê chuột)
        void FitToolbar()
        {
            int need = bar.Padding.Horizontal;
            foreach (Control c in bar.Controls)
            {
                var b = c as FButton;
                need += (b != null && b.Compactable ? b.FullWidth : c.Width) + c.Margin.Horizontal;
            }
            bool compact = need > bar.ClientSize.Width;
            foreach (var b in bar.Controls.OfType<FButton>()) if (b.Compactable) b.ShowText = !compact;
        }

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
                ? "Kéo file từ danh sách bên trái hoặc từ File Explorer vào vùng bên phải. Kéo trang để sắp xếp, giữ Ctrl khi thả để sao chép."
                : ws.Docs.Count + " tài liệu  ·  " + pages + " trang" + (sel ? "  ·  đang chọn " + ws.Selected.Count + " trang" : "") +
                  "      Kéo trang để sắp xếp hoặc chuyển sang tài liệu khác, giữ Ctrl để sao chép.";
        }

        void PreviewSelected()
        {
            var s = ws.OrderedSelection();
            if (s.Count > 0) Preview(s[0]);
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
                    statusBar.Refresh();
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
        readonly NavSegment nav;
        readonly CreatorView creator;
        readonly ToolboxView toolbox;
        readonly FButton themeBtn;

        static int S(int v) { return Theme.S(v); }

        public MainForm(string[] args)
        {
            Text = "PdfTool - Bộ công cụ PDF";
            Font = Theme.Ui(10f);
            Size = new Size(S(1400), S(880));
            MinimumSize = new Size(S(1100), S(680));
            StartPosition = FormStartPosition.CenterScreen;
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }
            if (Settings.Get("maximized", "1") == "1") WindowState = FormWindowState.Maximized;

            var content = new Panel { Dock = DockStyle.Fill };
            Theme.Bg(content, Role.Window);
            creator = new CreatorView(Theme.Scale) { Dock = DockStyle.Fill };
            toolbox = new ToolboxView(Theme.Scale) { Dock = DockStyle.Fill, Visible = false };
            content.Controls.Add(creator);
            content.Controls.Add(toolbox);

            // Đầu trang: logo, tên, chuyển trang, chọn giao diện sáng/tối
            var header = new TableLayoutPanel { Dock = DockStyle.Top, Height = S(60), ColumnCount = 6, RowCount = 1, Padding = new Padding(S(18), 0, S(14), 0) };
            Theme.Bg(header, Role.Window);
            for (int i = 0; i < 4; i++) header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            header.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var logo = new PictureBox { Size = new Size(S(28), S(28)), SizeMode = PictureBoxSizeMode.Zoom, Anchor = AnchorStyles.Left, Margin = new Padding(0, 0, S(10), 0) };
            try { logo.Image = new Icon(Icon, 32, 32).ToBitmap(); } catch { }
            var title = new Label { Text = "PdfTool", AutoSize = true, Font = Theme.UiSemi(13f), Anchor = AnchorStyles.Left, Margin = new Padding(0) };
            var sub = new Label { Text = "Bộ công cụ PDF", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(S(8), S(3), S(28), 0) };
            Theme.Fg(sub, TextRole.Muted);
            nav = new NavSegment { Anchor = AnchorStyles.Left, Margin = new Padding(0) };
            nav.Add("\uE70F", "Soạn PDF");
            nav.Add("\uE90F", "Hộp công cụ");
            nav.SelectedIndexChanged += (s, e) => { creator.Visible = nav.SelectedIndex == 0; toolbox.Visible = nav.SelectedIndex == 1; };
            themeBtn = new FButton(null, "\uE706", ButtonKind.Subtle) { Anchor = AnchorStyles.Right, Margin = new Padding(0) };
            themeBtn.Click += (s, e) => ShowThemeMenu();
            header.Controls.Add(logo, 0, 0);
            header.Controls.Add(title, 1, 0);
            header.Controls.Add(sub, 2, 0);
            header.Controls.Add(nav, 3, 0);
            header.Controls.Add(themeBtn, 5, 0);
            var headerLine = new Panel { Dock = DockStyle.Top, Height = 1 };
            Theme.Bg(headerLine, Role.Divider);

            Controls.Add(content);
            Controls.Add(headerLine);
            Controls.Add(header);
            if (Settings.Get("tab", "0") == "1") nav.SelectedIndex = 1;

            creator.SendToToolbox = path => { nav.SelectedIndex = 1; toolbox.AddPdf(path); };
            Theme.Changed += OnThemeChanged;
            UpdateThemeButton();
            Theme.Apply(this);

            Load += (s, e) =>
            {
                if (args.Length > 0) { nav.SelectedIndex = 0; creator.OpenFiles(args, null, -1); }
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
                Theme.Changed -= OnThemeChanged;
                toolbox.SaveSettings();
                creator.SaveSettings();
                Settings.Set("tab", nav.SelectedIndex.ToString());
                Settings.Set("maximized", WindowState == FormWindowState.Maximized ? "1" : "0");
                Settings.Save();
                creator.Cleanup();
            };
        }

        void OnThemeChanged()
        {
            if (IsDisposed) return;
            if (InvokeRequired) { BeginInvoke(new Action(OnThemeChanged)); return; }
            Theme.Apply(this);
            UpdateThemeButton();
            Invalidate(true);
        }

        void UpdateThemeButton()
        {
            themeBtn.Glyph = Theme.P.Dark ? "\uE708" : "\uE706";
            string mode = Theme.Mode == ThemeMode.System ? "theo Windows" : Theme.Mode == ThemeMode.Light ? "sáng" : "tối";
            Theme.Tip(themeBtn, "Giao diện: " + mode);
        }

        void ShowThemeMenu()
        {
            var m = Theme.Menu();
            m.ShowCheckMargin = true;
            Theme.Item(m, "Theo Windows", "\uE713", (s, e) => Theme.SetMode(ThemeMode.System)).Checked = Theme.Mode == ThemeMode.System;
            Theme.Item(m, "Sáng", "\uE706", (s, e) => Theme.SetMode(ThemeMode.Light)).Checked = Theme.Mode == ThemeMode.Light;
            Theme.Item(m, "Tối", "\uE708", (s, e) => Theme.SetMode(ThemeMode.Dark)).Checked = Theme.Mode == ThemeMode.Dark;
            m.Closed += (s, e) => BeginInvoke(new Action(m.Dispose));
            m.Show(themeBtn, new Point(themeBtn.Width - m.GetPreferredSize(Size.Empty).Width, themeBtn.Height + S(4)));
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
            Theme.Init();
            Application.Run(new MainForm(args));
        }
    }
}
