using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media.Imaging;

namespace OCRBulkAdd.Utils
{
    internal static class ImageInput
    {
        private static readonly HashSet<string> ImageExtensions =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp" };


        public static bool TryGetFromClipboard(out byte[] bytes, int retries = 3, int retryDelayMs = 50)
        {
            // get clipboard data
            if (TryGetClipboardBitmap(out var bmp, retries, retryDelayMs))
            {
                bytes = ImageHelper.BitmapSourceToPngBytes(bmp!);
                return true;
            }

            if (TryGetClipboardImageFile(out var path, retries, retryDelayMs))
            {
                bytes = File.ReadAllBytes(path);
                return true;
            }

            bytes = Array.Empty<byte>();
            return false;
        }

        public static bool HasImageData(IDataObject data)
        {
            if (data == null) return false;

            if (data.GetDataPresent(DataFormats.Bitmap))
                return true;

            if (data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = data.GetData(DataFormats.FileDrop) as string[];
                return files != null && files.Any(IsLikelyImageFile);
            }

            return false;
        }

        public static bool TryGetFromDataObject(IDataObject data, out byte[] bytes)
        {
            // get drag-n-drop data
            bytes = Array.Empty<byte>();
            if (data == null) return false;

            if (data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = data.GetData(DataFormats.FileDrop) as string[];
                string? file = files?.FirstOrDefault(IsLikelyImageFile);

                if (!string.IsNullOrWhiteSpace(file) && File.Exists(file))
                {
                    bytes = File.ReadAllBytes(file);
                    return true;
                }
            }

            if (data.GetDataPresent(DataFormats.Bitmap))
            {
                var bmp = data.GetData(DataFormats.Bitmap) as BitmapSource;
                if (bmp != null)
                {
                    bytes = ImageHelper.BitmapSourceToPngBytes(bmp);
                    return true;
                }
            }

            return false;
        }

        private static bool TryGetClipboardBitmap(out BitmapSource? bitmap, int retries, int retryDelayMs)
        {
            bitmap = null;

            for (int attempt = 0; attempt < retries; attempt++)
            {
                try
                {
                    if (Clipboard.ContainsImage())
                    {
                        var bmp = Clipboard.GetImage();
                        if (bmp != null)
                        {
                            bitmap = bmp;
                            return true;
                        }
                    }
                    return false;
                }
                catch (COMException)
                {
                    Thread.Sleep(retryDelayMs);
                }
            }

            return false;
        }

        private static bool TryGetClipboardImageFile(out string path, int retries, int retryDelayMs)
        {
            path = string.Empty;

            for (int attempt = 0; attempt < retries; attempt++)
            {
                try
                {
                    if (!Clipboard.ContainsFileDropList())
                        return false;

                    var files = Clipboard.GetFileDropList();
                    if (files.Count == 0)
                        return false;

                    string? file = files.Cast<string>().FirstOrDefault(IsLikelyImageFile);
                    if (!string.IsNullOrWhiteSpace(file))
                    {
                        path = file;
                        return true;
                    }

                    return false;
                }
                catch (COMException)
                {
                    Thread.Sleep(retryDelayMs);
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
    }
}