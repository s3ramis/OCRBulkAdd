using System.IO;
using System.Windows.Media.Imaging;

namespace OCRBulkAdd.Utils
{
    internal static class ImageHelper
    {
        public static byte[] BitmapSourceToPngBytes(BitmapSource source)
        {
            // convert bitmap to bytearray for tesseract to consume
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(source));

            using (var ms = new MemoryStream())
            {
                encoder.Save(ms);
                return ms.ToArray();
            }
        }

        public static BitmapImage BytesToBitmapImage(byte[] bytes)
        {
            // convert bytearray to image for preview box
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
    }
}