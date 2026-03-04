using System;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace OCRBulkAdd.Processing
{
    internal static class ImagePreprocessor
    {
        /// <summary>
        /// make the input imable as easily ocr-able as possible
        /// black text on white background
        ///
        ///  1. un-transparency
        ///  2. upscaling
        ///  3) convert to gray8
        ///  4) auto contrast
        ///  5) binarization (black/white)
        ///  6) add border
        /// </summary>
        public static BitmapSource PrepareForOcr(
            BitmapSource src,
            double scale = 2.0,
            int borderPx = 12,
            double lowCutPercent = 0.01,
            double highCutPercent = 0.99)
        {
            if (src == null) throw new ArgumentNullException(nameof(src));

            // 1) If PNGs have alpha, composite them on white so antialiased edges stay light.
            BitmapSource opaque = CompositeOnWhite(src);

            // 2) Upscale to improve recognition of punctuation.
            var scaled = new TransformedBitmap(opaque, new ScaleTransform(scale, scale));
            scaled.Freeze();

            // 3) Grayscale (8-bit).
            var gray = new FormatConvertedBitmap(scaled, PixelFormats.Gray8, null, 0);
            gray.Freeze();

            // 4) "Levels": stretch the histogram to boost contrast and brighten near-white backgrounds.
            var contrasted = AutoContrastGray8(gray, lowCutPercent, highCutPercent);

            // 5) Convert to pure black/white.
            var binary = BinarizeOtsuGray8(contrasted);

            // 6) Border for better segmentation.
            return AddBorderGray8(binary, borderPx);
        }

        private static BitmapSource CompositeOnWhite(BitmapSource src)
        {
            // Convert to BGRA to access alpha channel.
            var bgra = new FormatConvertedBitmap(src, PixelFormats.Bgra32, null, 0);
            bgra.Freeze();

            int w = bgra.PixelWidth;
            int h = bgra.PixelHeight;
            int stride = w * 4;

            byte[] px = new byte[stride * h];
            bgra.CopyPixels(px, stride, 0);

            // Alpha blend each pixel on white:
            // out = src*a + 255*(1-a)
            for (int i = 0; i < px.Length; i += 4)
            {
                byte b = px[i];
                byte g = px[i + 1];
                byte r = px[i + 2];
                byte a = px[i + 3];

                if (a != 255)
                {
                    int inv = 255 - a;
                    px[i]     = (byte)((b * a + 255 * inv + 127) / 255);
                    px[i + 1] = (byte)((g * a + 255 * inv + 127) / 255);
                    px[i + 2] = (byte)((r * a + 255 * inv + 127) / 255);
                    px[i + 3] = 255;
                }
                else
                {
                    px[i + 3] = 255;
                }
            }

            var wb = new WriteableBitmap(w, h, 96, 96, PixelFormats.Bgra32, null);
            wb.WritePixels(new Int32Rect(0, 0, w, h), px, stride, 0);
            wb.Freeze();
            return wb;
        }

        private static BitmapSource AutoContrastGray8(BitmapSource gray, double lowCutPercent, double highCutPercent)
        {
            if (gray.Format != PixelFormats.Gray8)
            {
                gray = new FormatConvertedBitmap(gray, PixelFormats.Gray8, null, 0);
                gray.Freeze();
            }

            int w = gray.PixelWidth;
            int h = gray.PixelHeight;
            int stride = w;

            byte[] px = new byte[stride * h];
            gray.CopyPixels(px, stride, 0);

            // Build histogram
            int[] hist = new int[256];
            for (int i = 0; i < px.Length; i++)
                hist[px[i]]++;

            int total = px.Length;
            int lowTarget = (int)(total * lowCutPercent);
            int highTarget = (int)(total * highCutPercent);

            // Find low/high percentiles (ignore outliers)
            int cum = 0;
            int low = 0;
            for (int i = 0; i < 256; i++)
            {
                cum += hist[i];
                if (cum >= lowTarget) { low = i; break; }
            }

            cum = 0;
            int high = 255;
            for (int i = 0; i < 256; i++)
            {
                cum += hist[i];
                if (cum >= highTarget) { high = i; break; }
            }

            // If no meaningful dynamic range, skip.
            if (high <= low + 1)
                return gray;

            int range = high - low;

            // Linear stretch: [low..high] -> [0..255]
            for (int i = 0; i < px.Length; i++)
            {
                int v = (px[i] - low) * 255 / range;
                if (v < 0) v = 0;
                if (v > 255) v = 255;
                px[i] = (byte)v;
            }

            var bmp = BitmapSource.Create(w, h, 96, 96, PixelFormats.Gray8, null, px, stride);
            bmp.Freeze();
            return bmp;
        }

        private static BitmapSource BinarizeOtsuGray8(BitmapSource gray)
        {
            if (gray.Format != PixelFormats.Gray8)
            {
                gray = new FormatConvertedBitmap(gray, PixelFormats.Gray8, null, 0);
                gray.Freeze();
            }

            int w = gray.PixelWidth;
            int h = gray.PixelHeight;
            int stride = w;

            byte[] px = new byte[stride * h];
            gray.CopyPixels(px, stride, 0);

            // Histogram
            int[] hist = new int[256];
            for (int i = 0; i < px.Length; i++)
                hist[px[i]]++;

            // Otsu: maximize between-class variance
            int total = px.Length;
            double sum = 0;
            for (int t = 0; t < 256; t++)
                sum += t * hist[t];

            double sumB = 0;
            int wB = 0;
            double maxVar = -1;
            int threshold = 128;

            for (int t = 0; t < 256; t++)
            {
                wB += hist[t];
                if (wB == 0) continue;

                int wF = total - wB;
                if (wF == 0) break;

                sumB += t * hist[t];
                double mB = sumB / wB;
                double mF = (sum - sumB) / wF;

                double varBetween = (double)wB * wF * (mB - mF) * (mB - mF);
                if (varBetween > maxVar)
                {
                    maxVar = varBetween;
                    threshold = t;
                }
            }

            // Apply threshold -> binary
            byte[] bin = new byte[px.Length];
            int black = 0, white = 0;

            for (int i = 0; i < px.Length; i++)
            {
                byte v = (px[i] <= threshold) ? (byte)0 : (byte)255;
                bin[i] = v;
                if (v == 0) black++; else white++;
            }

            // If background is mostly black, invert to "black text on white"
            if (black > white)
            {
                for (int i = 0; i < bin.Length; i++)
                    bin[i] = (byte)(255 - bin[i]);
            }

            var bmp = BitmapSource.Create(w, h, 96, 96, PixelFormats.Gray8, null, bin, stride);
            bmp.Freeze();
            return bmp;
        }

        private static BitmapSource AddBorderGray8(BitmapSource gray, int borderPx)
        {
            if (gray.Format != PixelFormats.Gray8)
            {
                gray = new FormatConvertedBitmap(gray, PixelFormats.Gray8, null, 0);
                gray.Freeze();
            }

            int w = gray.PixelWidth;
            int h = gray.PixelHeight;
            int stride = w;

            byte[] srcPx = new byte[stride * h];
            gray.CopyPixels(srcPx, stride, 0);

            int newW = w + borderPx * 2;
            int newH = h + borderPx * 2;
            int newStride = newW;

            // Fill output with white.
            byte[] outPx = new byte[newStride * newH];
            Array.Fill(outPx, (byte)255);

            // Copy source into the center.
            for (int y = 0; y < h; y++)
            {
                Buffer.BlockCopy(
                    srcPx, y * stride,
                    outPx, (y + borderPx) * newStride + borderPx,
                    w);
            }

            var bmp = BitmapSource.Create(newW, newH, 96, 96, PixelFormats.Gray8, null, outPx, newStride);
            bmp.Freeze();
            return bmp;
        }
    }
}