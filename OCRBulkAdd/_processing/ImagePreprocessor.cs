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
        public static BitmapSource PrepareForOcr(BitmapSource src, PreprocessingSettings settings)
        {
            if (src == null) throw new ArgumentNullException(nameof(src));
            if (settings == null) throw new ArgumentNullException(nameof(settings));

            // If disabled, return as-is (no transformation).
            if (!settings.Enabled)
                return src;

            // 1) Composite transparency on white so antialiased edges don't become gray/dark.
            BitmapSource opaque = CompositeOnWhite(src);

            // 2) Upscale: small punctuation (minus/comma) gets more pixels -> higher hit-rate.
            BitmapSource scaled = opaque;
            if (Math.Abs(settings.Scale - 1.0) > 0.0001)
            {
                var t = new TransformedBitmap(opaque, new ScaleTransform(settings.Scale, settings.Scale));
                t.Freeze();
                scaled = t;
            }

            // 3) Convert to Gray8 (needed for histogram + Otsu).
            var gray = new FormatConvertedBitmap(scaled, PixelFormats.Gray8, null, 0);
            gray.Freeze();

            BitmapSource working = gray;

            // 4) Auto-contrast: stretches histogram, often turning "off-white" background into white.
            if (settings.EnableAutoContrast)
                working = AutoContrastGray8(working, settings.LowCutPercent, settings.HighCutPercent);

            // 5) Binarize: converts to pure black/white; helps when background is uneven.
            if (settings.EnableBinarization)
                working = BinarizeOtsuGray8(working, autoInvert: settings.AutoInvert);

            // 6) Add a clean white border: helps layout analysis and avoids cropping at edges.
            if (settings.BorderPx > 0)
                working = AddBorderGray8(working, settings.BorderPx);

            return working;
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

        private static BitmapSource BinarizeOtsuGray8(BitmapSource gray, bool autoInvert)
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
            if (autoInvert && black > white)
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