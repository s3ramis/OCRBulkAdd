using System;
using System.ComponentModel;

namespace OCRBulkAdd.Processing
{
    public enum PreprocessingPreset
    {
        Default,
        Aggressive,
        NoBinarization,
        Disabled
    }

    public sealed class PreprocessingSettings : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        private bool _enabled = true;
        public bool Enabled
        {
            get => _enabled;
            set { if (_enabled != value) { _enabled = value; OnChanged(nameof(Enabled)); } }
        }

        private bool _enableAutoContrast = true;
        public bool EnableAutoContrast
        {
            get => _enableAutoContrast;
            set { if (_enableAutoContrast != value) { _enableAutoContrast = value; OnChanged(nameof(EnableAutoContrast)); } }
        }

        private bool _enableBinarization = true;
        public bool EnableBinarization
        {
            get => _enableBinarization;
            set { if (_enableBinarization != value) { _enableBinarization = value; OnChanged(nameof(EnableBinarization)); } }
        }

        private bool _autoInvert = true;
        public bool AutoInvert
        {
            get => _autoInvert;
            set { if (_autoInvert != value) { _autoInvert = value; OnChanged(nameof(AutoInvert)); } }
        }

        private double _scale = 2.0;
        public double Scale
        {
            get => _scale;
            set
            {
                double v = Math.Clamp(value, 1.0, 4.0);
                if (Math.Abs(_scale - v) > 0.0001)
                {
                    _scale = v;
                    OnChanged(nameof(Scale));
                }
            }
        }

        private int _borderPx = 12;
        public int BorderPx
        {
            get => _borderPx;
            set
            {
                int v = Math.Clamp(value, 0, 80);
                if (_borderPx != v)
                {
                    _borderPx = v;
                    OnChanged(nameof(BorderPx));
                }
            }
        }

        private double _lowCutPercent = 0.01;
        public double LowCutPercent
        {
            get => _lowCutPercent;
            set
            {
                double v = Math.Clamp(value, 0.0, 0.2);
                if (Math.Abs(_lowCutPercent - v) > 0.0001)
                {
                    _lowCutPercent = v;
                    if (_highCutPercent <= _lowCutPercent) HighCutPercent = Math.Min(1.0, _lowCutPercent + 0.01);
                    OnChanged(nameof(LowCutPercent));
                }
            }
        }

        private double _highCutPercent = 0.99;
        public double HighCutPercent
        {
            get => _highCutPercent;
            set
            {
                double v = Math.Clamp(value, 0.0, 1.0);
                if (Math.Abs(_highCutPercent - v) > 0.0001)
                {
                    _highCutPercent = v;
                    if (_highCutPercent <= _lowCutPercent) LowCutPercent = Math.Max(0.0, _highCutPercent - 0.01);
                    OnChanged(nameof(HighCutPercent));
                }
            }
        }

        public void ApplyPreset(PreprocessingPreset preset)
        {
            switch (preset)
            {
                case PreprocessingPreset.Default:
                    Enabled = true;
                    Scale = 2.0;
                    BorderPx = 12;
                    EnableAutoContrast = true;
                    LowCutPercent = 0.01;
                    HighCutPercent = 0.99;
                    EnableBinarization = true;
                    AutoInvert = true;
                    break;

                case PreprocessingPreset.Aggressive:
                    Enabled = true;
                    Scale = 2.5;
                    BorderPx = 12;
                    EnableAutoContrast = true;
                    LowCutPercent = 0.02;
                    HighCutPercent = 0.985;
                    EnableBinarization = true;
                    AutoInvert = true;
                    break;

                case PreprocessingPreset.NoBinarization:
                    Enabled = true;
                    Scale = 2.0;
                    BorderPx = 12;
                    EnableAutoContrast = true;
                    LowCutPercent = 0.01;
                    HighCutPercent = 0.99;
                    EnableBinarization = false;
                    AutoInvert = true;
                    break;

                case PreprocessingPreset.Disabled:
                    Enabled = false;
                    Scale = 1.0;
                    BorderPx = 0;
                    EnableAutoContrast = false;
                    LowCutPercent = 0.01;
                    HighCutPercent = 0.99;
                    EnableBinarization = false;
                    AutoInvert = true;
                    break;
            }
        }
    }
}