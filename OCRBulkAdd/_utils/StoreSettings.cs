using System.IO;
using System.Text.Json;
using OCRBulkAdd.Processing;

namespace OCRBulkAdd.Utils
{
    internal static class SettingsStore
    {
        private const int CurrentVersion = 1;

        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            WriteIndented = true,
            AllowTrailingCommas = true,
            ReadCommentHandling = JsonCommentHandling.Skip
        };

        private static string SettingsDir =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OCRBulkAdd");

        private static string SettingsFilePath =>
            Path.Combine(SettingsDir, "settings.json");

        private sealed class PersistedSettings
        {
            public int Version { get; set; } = CurrentVersion;
            public PreprocessSnapshot Preprocess { get; set; } = new();
        }

        private sealed class PreprocessSnapshot
        {
            public bool Enabled { get; set; }
            public bool EnableAutoContrast { get; set; }
            public bool EnableBinarization { get; set; }
            public bool AutoInvert { get; set; }
            public double Scale { get; set; }
            public int BorderPx { get; set; }
            public double LowCutPercent { get; set; }
            public double HighCutPercent { get; set; }
        }

        public static void Save(PreprocessingSettings preprocess)
        {
            try
            {
                Directory.CreateDirectory(SettingsDir);

                var model = new PersistedSettings
                {
                    Version = CurrentVersion,
                    Preprocess = new PreprocessSnapshot
                    {
                        Enabled = preprocess.Enabled,
                        EnableAutoContrast = preprocess.EnableAutoContrast,
                        EnableBinarization = preprocess.EnableBinarization,
                        AutoInvert = preprocess.AutoInvert,
                        Scale = preprocess.Scale,
                        BorderPx = preprocess.BorderPx,
                        LowCutPercent = preprocess.LowCutPercent,
                        HighCutPercent = preprocess.HighCutPercent
                    }
                };

                var json = JsonSerializer.Serialize(model, JsonOptions);
                File.WriteAllText(SettingsFilePath, json);
            }
            catch
            {
                
            }
        }

        public static bool TryLoadInto(PreprocessingSettings target)
        {
            try
            {
                if (!File.Exists(SettingsFilePath))
                    return false;

                var json = File.ReadAllText(SettingsFilePath);
                var model = JsonSerializer.Deserialize<PersistedSettings>(json, JsonOptions);

                if (model?.Preprocess == null)
                    return false;

                target.Enabled = model.Preprocess.Enabled;
                target.EnableAutoContrast = model.Preprocess.EnableAutoContrast;
                target.EnableBinarization = model.Preprocess.EnableBinarization;
                target.AutoInvert = model.Preprocess.AutoInvert;
                target.Scale = model.Preprocess.Scale;
                target.BorderPx = model.Preprocess.BorderPx;
                target.LowCutPercent = model.Preprocess.LowCutPercent;
                target.HighCutPercent = model.Preprocess.HighCutPercent;

                return true;
            }
            catch
            {
                return false;
            }
        }

        public static void Delete()
        {
            try
            {
                if (File.Exists(SettingsFilePath))
                    File.Delete(SettingsFilePath);
            }
            catch { }
        }
    }
}