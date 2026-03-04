using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using TesseractOCR;
using TesseractOCR.Enums;
using OcrLanguage = TesseractOCR.Enums.Language;

namespace OCRBulkAdd.Services
{
    internal sealed class OcrService
    {
        public string TessdataPath { get; }

        public EngineMode EngineMode { get; set; } = EngineMode.LstmOnly;
        public PageSegMode PageSegMode { get; set; } = PageSegMode.SparseText;

        // debug, tesseract saves image its trying to ocr
        public bool WriteDebugImages { get; set; } = false;

        // recommended dpi for ocr (if not already set by input file)
        public int UserDefinedDpi { get; set; } = 300;

        public OcrService(string? tessdataPath = null)
        {
            TessdataPath = tessdataPath ?? TessdataPathResolver.Resolve();
        }

        public Task<string> RecognizeAsync(byte[] imageBytes, CancellationToken ct = default)
        {
            if (imageBytes == null) throw new ArgumentNullException(nameof(imageBytes));
            if (imageBytes.Length == 0) return Task.FromResult(string.Empty);

            return Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();

                if (!Directory.Exists(TessdataPath))
                    throw new DirectoryNotFoundException("missing tessdata folder: " + TessdataPath);

                List<OcrLanguage> languages = DetectLanguages(TessdataPath);

                // disable dictionaries since we primarily care about the numbers
                var initialValues = new Dictionary<string, object>
                {
                    ["load_system_dawg"] = 0,
                    ["load_freq_dawg"] = 0,
                    ["user_defined_dpi"] = UserDefinedDpi
                };

                using (var engine = new Engine(TessdataPath, languages, EngineMode, initialValues: initialValues))
                {
                    engine.DefaultPageSegMode = PageSegMode;

                    if (WriteDebugImages)
                        engine.SetVariable("tessedit_write_images", true);

                    using (var img = TesseractOCR.Pix.Image.LoadFromMemory(imageBytes))
                    using (var page = engine.Process(img, PageSegMode))
                    {
                        return page.Text ?? string.Empty;
                    }
                }
            }, ct);
        }

        private static List<OcrLanguage> DetectLanguages(string tessdataPath)
        {
            var langs = new List<OcrLanguage>();

            string eng = Path.Combine(tessdataPath, "eng.traineddata");
            string deu = Path.Combine(tessdataPath, "deu.traineddata");

            if (File.Exists(eng)) langs.Add(OcrLanguage.English);
            if (File.Exists(deu)) langs.Add(OcrLanguage.German);

            if (langs.Count == 0)
                throw new FileNotFoundException("no traineddata found. put e.g. eng.traineddata into: " + tessdataPath);

            return langs;
        }
    }

    internal static class TessdataPathResolver
    {
        public static string Resolve()
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
    }
}