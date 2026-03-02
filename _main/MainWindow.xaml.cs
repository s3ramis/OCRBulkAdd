using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Input;
using TesseractOCR;
using TesseractOCR.Enums;

using OcrLanguage = TesseractOCR.Enums.Language;

namespace OCRBulkAdd
{
    public partial class MainWindow : Window
    {
        private byte[] _currentImageBytes = Array.Empty<byte>();

        private readonly SemaphoreSlim _ocrLock = new SemaphoreSlim(1, 1);

        private static readonly HashSet<string> ImageExtensions =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp" };

        private static readonly string TessdataPath = ResolveTessdataPath();

        public MainWindow()
        {
            InitializeComponent();
            UpdateHintVisibility();
        }

        private static string ResolveTessdataPath()
        {
            string baseDir = AppContext.BaseDirectory;

            string candidate = Path.Combine(baseDir, "tessdata");
            if (Directory.Exists(candidate)) return candidate;

            candidate = Path.Combine(Directory.GetCurrentDirectory(), "tessdata");
            if (Directory.Exists(candidate)) return candidate;

            candidate = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "tessdata"));
            if (Directory.Exists(candidate)) return candidate;

            return Path.Combine(baseDir, "tessdata");
        }

        private async void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.V) return;
            if ((Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control) return;

            if (TryGetClipboardImageBytes(out var bytes) || TryGetClipboardImageFileBytes(out bytes))
            {
                SetImageBytes(bytes);
                e.Handled = true;
                await RunOcrAsync();
            }
        }

        private static bool TryGetClipboardImageBytes(out byte[] bytes)
        {
            bytes = Array.Empty<byte>();

            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    if (Clipboard.ContainsImage())
                    {
                        var bmp = Clipboard.GetImage();
                        if (bmp != null)
                        {
                            bmp = PrepareForOcr(bmp);
                            bytes = BitmapSourceToPngBytes(bmp);
                            return true;
                        }
                    }
                    return false;
                }
                catch (COMException)
                {
                    Thread.Sleep(50);
                }
            }

            return false;
        }

        private static bool TryGetClipboardImageFileBytes(out byte[] bytes)
        {
            bytes = Array.Empty<byte>();

            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    if (!Clipboard.ContainsFileDropList())
                        return false;

                    var files = Clipboard.GetFileDropList();
                    if (files.Count == 0)
                        return false;

                    string? file = files.Cast<string>().FirstOrDefault(IsLikelyImageFile);

                    if (string.IsNullOrWhiteSpace(file) || !File.Exists(file))
                        return false;

                    bytes = File.ReadAllBytes(file);
                    return true;
                }
                catch (COMException)
                {
                    Thread.Sleep(50);
                }
            }

            return false;
        }

        private void Window_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = HasImageData(e.Data) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        private async void Window_Drop(object sender, DragEventArgs e)
        {
            if (TryGetImageBytes(e.Data, out var bytes))
            {
                SetImageBytes(bytes);
                await RunOcrAsync();
            }
        }

        private static bool HasImageData(IDataObject data)
        {
            if (data.GetDataPresent(DataFormats.Bitmap))
                return true;

            if (data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = data.GetData(DataFormats.FileDrop) as string[];
                return files != null && files.Any(IsLikelyImageFile);
            }

            return false;
        }

        private static bool TryGetImageBytes(IDataObject data, out byte[] bytes)
        {
            bytes = Array.Empty<byte>();

            if (data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = data.GetData(DataFormats.FileDrop) as string[];
                if (files != null)
                {
                    string? file = files.FirstOrDefault(IsLikelyImageFile);

                    if (!string.IsNullOrWhiteSpace(file) && File.Exists(file))
                    {
                        bytes = File.ReadAllBytes(file);
                        return true;
                    }
                }
            }

            if (data.GetDataPresent(DataFormats.Bitmap))
            {
                var bmp = data.GetData(DataFormats.Bitmap) as BitmapSource;
                if (bmp != null)
                {
                    bytes = BitmapSourceToPngBytes(bmp);
                    return true;
                }
            }

            return false;
        }

        private static bool IsLikelyImageFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;
            string ext = Path.GetExtension(path);
            return !string.IsNullOrWhiteSpace(ext) && ImageExtensions.Contains(ext);
        }

        private void SetImageBytes(byte[] bytes)
        {
            if (bytes.Length == 0)
                return;

            try
            {
                BitmapImage preview = BytesToBitmapImage(bytes);
                PreviewImage.Source = preview;

                _currentImageBytes = BitmapSourceToPngBytes(preview);
            }
            catch
            {
                PreviewImage.Source = null;
                _currentImageBytes = bytes; // fallback
            }

            OcrTextBox.Clear();
            StatusText.Text = "image loaded.";
            SumResultText.Text = "sum: -";
            UpdateHintVisibility();
        }

        private void UpdateHintVisibility()
        {
            DropHint.Visibility = PreviewImage.Source == null ? Visibility.Visible : Visibility.Collapsed;
        }

        private async void RunOcr_Click(object sender, RoutedEventArgs e)
        {
            await RunOcrAsync();
        }

        private void Sum_Click(object sender, RoutedEventArgs e)
        {
            var (sum, count) = SumNumbersFromText(OcrTextBox.Text);
            SumResultText.Text = $"sum: {sum.ToString("N", CultureInfo.CurrentCulture)}    (numbers found: {count})";
        }

        private async Task RunOcrAsync()
        {
            if (_currentImageBytes.Length == 0)
            {
                StatusText.Text = "no image yet (paste or drop one).";
                return;
            }

            await _ocrLock.WaitAsync();
            try
            {
                StatusText.Text = "running OCR...";
                byte[] bytes = _currentImageBytes;

                string text = await Task.Run(() =>
                {
                    if (!Directory.Exists(TessdataPath))
                        throw new DirectoryNotFoundException("missing tessdata folder: " + TessdataPath);

                    var languages = new List<OcrLanguage>();

                    string eng = Path.Combine(TessdataPath, "eng.traineddata");
                    string deu = Path.Combine(TessdataPath, "deu.traineddata");

                    if (File.Exists(eng)) languages.Add(OcrLanguage.English);
                    if (File.Exists(deu)) languages.Add(OcrLanguage.German);

                    if (languages.Count == 0)
                        throw new FileNotFoundException("no traineddata found. put e.g. traineddata into: " + TessdataPath);

                    var initialValues = new Dictionary<string, object>
                    {
                        ["load_system_dawg"] = 0,
                        ["load_freq_dawg"] = 0,
                        ["user_defined_dpi"] = 30
                    };

                    using (var engine = new Engine(TessdataPath, languages, EngineMode.LstmOnly, initialValues: initialValues))
                    {
                        engine.DefaultPageSegMode = PageSegMode.SparseText;
                        //engine.SetVariable("tessedit_write_images", true);
                        // engine.SetVariable("tessedit_char_whitelist", "0123456789.,-");

                        using (var img = TesseractOCR.Pix.Image.LoadFromMemory(bytes))
                        using (var page = engine.Process(img, PageSegMode.SparseText))
                        {
                            return page.Text ?? string.Empty;
                        }
                    }
                });

                OcrTextBox.Text = NormalizeOcrText(text);
                StatusText.Text = "OCR done.";
            }
            catch (Exception ex)
            {
                StatusText.Text = "OCR error: " + ex.Message;
            }
            finally
            {
                _ocrLock.Release();
            }
        }

        private static (decimal Sum, int Count) SumNumbersFromText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return (0m, 0);

            var matches = Regex.Matches(
                text,
                @"(?:[+-][ \u00A0\u202F]*)?(?:\d{1,3}(?:[., \u00A0\u202F]\d{3})+|\d+)(?:[.,]\d+)?");

            decimal sum = 0m;
            int count = 0;

            var cultures = new[]
            {
                CultureInfo.GetCultureInfo("de-DE"),
                CultureInfo.GetCultureInfo("en-US"),
                CultureInfo.InvariantCulture
            };

            foreach (Match m in matches)
            {
                string token = m.Value.Trim();

                token = token.Replace(" ", "")
                            .Replace("\u00A0", "")   // NBSP
                            .Replace("\u202F", "");  // narrow NBSP

                foreach (var culture in cultures)
                {
                    if (decimal.TryParse(token,
                        NumberStyles.Number | NumberStyles.AllowLeadingSign,
                        culture,
                        out decimal value))
                    {
                        sum += value;
                        count++;
                        break;
                    }
                }
            }

            return (sum, count);
        }

        private static byte[] BitmapSourceToPngBytes(BitmapSource source)
        {
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(source));

            using (var ms = new MemoryStream())
            {
                encoder.Save(ms);
                return ms.ToArray();
            }
        }

        private static BitmapImage BytesToBitmapImage(byte[] bytes)
        {
            using (var ms = new MemoryStream(bytes))
            {
                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.StreamSource = ms;
                bmp.EndInit();
                bmp.Freeze();
                return bmp;
            }
        }

        private static BitmapSource PrepareForOcr(BitmapSource src)
        {
            var rgb = new FormatConvertedBitmap(src, PixelFormats.Bgr24, null, 0);
            rgb.Freeze();

            var scaled = new TransformedBitmap(rgb, new ScaleTransform(2.0, 2.0));
            scaled.Freeze();

            return AddBorder(scaled, 10);
        }

        private static BitmapSource AddBorder(BitmapSource src, int borderPx)
        {
            int w = src.PixelWidth + borderPx * 2;
            int h = src.PixelHeight + borderPx * 2;

            var dv = new DrawingVisual();
            using (var dc = dv.RenderOpen())
            {
                dc.DrawRectangle(Brushes.White, null, new System.Windows.Rect(0, 0, w, h));
                dc.DrawImage(src, new System.Windows.Rect(borderPx, borderPx, src.PixelWidth, src.PixelHeight));
            }

            var rtb = new RenderTargetBitmap(w, h, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(dv);
            rtb.Freeze();
            return rtb;
        }

        private static string NormalizeOcrText(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return string.Empty;

            s = s.Replace("\r\n", "\n").Replace('\r', '\n');

            var lines = s.Split('\n')
                        .Select(l => l.Trim())
                        .Where(l => !string.IsNullOrWhiteSpace(l));

            return string.Join(Environment.NewLine, lines).Trim();
        }
    }
}