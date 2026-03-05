using System.ComponentModel;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using OCRBulkAdd.Processing;
using OCRBulkAdd.Services;
using OCRBulkAdd.Utils;

namespace OCRBulkAdd
{
    public partial class MainWindow : Window
    {
        public PreprocessingSettings Preprocess { get; } = new PreprocessingSettings();

        private byte[] _originalImageBytes = Array.Empty<byte>();
        private byte[] _currentOcrImageBytes = Array.Empty<byte>();

        private readonly SemaphoreSlim _ocrLock = new SemaphoreSlim(1, 1);
        private readonly OcrService _ocrService = new OcrService();

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;

            SettingsStore.TryLoadInto(Preprocess);

            Closing += MainWindow_Closing;

            UpdateHintVisibility();
        }

        private async void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.V) return;
            if ((Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control) return;

            if (ImageInput.TryGetFromClipboard(out var bytes))
            {
                SetImage(bytes);
                e.Handled = true;
                await RunOcrAsync();
            }
        }

        private void Window_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = ImageInput.HasImageData(e.Data) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        private async void Window_Drop(object sender, DragEventArgs e)
        {
            if (ImageInput.TryGetFromDataObject(e.Data, out var bytes))
            {
                SetImage(bytes);
                await RunOcrAsync();
            }
        }

        private void SetImage(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0) return;

            _originalImageBytes = bytes;

            var preview = ImageHelper.BytesToBitmapImage(bytes);
            PreviewImage.Source = preview;

            RebuildOcrBytesFromSettings();

            OcrPreviewTextBox.Clear();
            NumbersTextBox.Clear();
            SumResultBox.Clear();
            SumMetaText.Text = "";

            StatusText.Text = "image loaded.";
            UpdateHintVisibility();
        }

        private void RebuildOcrBytesFromSettings()
        {
            if (_originalImageBytes.Length == 0) return;

            // always start preprocessing from original image
            var src = ImageHelper.BytesToBitmapImage(_originalImageBytes);
            var prepared = ImagePreprocessor.PrepareForOcr(src, Preprocess);
            _currentOcrImageBytes = ImageHelper.BitmapSourceToPngBytes(prepared);
        }

        private void UpdateHintVisibility()
        {
            DropHint.Visibility = PreviewImage.Source == null ? Visibility.Visible : Visibility.Collapsed;
        }

        private async void RunOcr_Click(object sender, RoutedEventArgs e)
        {
            await RunOcrAsync();
        }

        private async void ApplyPreprocessAndOcr_Click(object sender, RoutedEventArgs e)
        {
            await RunOcrAsync();
        }

        private void PreprocessPreset_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            switch (PreprocessPresetCombo.SelectedIndex)
            {
                case 0: Preprocess.ApplyPreset(PreprocessingPreset.Default); break;
                case 1: Preprocess.ApplyPreset(PreprocessingPreset.Aggressive); break;
                case 2: Preprocess.ApplyPreset(PreprocessingPreset.NoBinarization); break;
                case 3: Preprocess.ApplyPreset(PreprocessingPreset.Disabled); break;
            }
        }

        private async Task RunOcrAsync()
        {
            if (_originalImageBytes.Length == 0)
            {
                StatusText.Text = "no image yet (paste or drop one).";
                return;
            }

            await _ocrLock.WaitAsync();
            try
            {
                StatusText.Text = "running OCR...";

                // apply ocr settings from ui before ocr process starts
                RebuildOcrBytesFromSettings();

                string raw = await _ocrService.RecognizeAsync(_currentOcrImageBytes);

                string previewText = OcrTextNormalizer.Normalize(raw);
                OcrPreviewTextBox.Text = previewText;

                NumbersTextBox.Text = NumberExtractor.ExtractNormalizedNumbersText(previewText);

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

        private void Sum_Click(object sender, RoutedEventArgs e)
        {
            var (sum, count) = NumberExtractor.SumFromText(NumbersTextBox.Text);

            SumResultBox.Text = sum.ToString("0.00", CultureInfo.InvariantCulture);
            SumMetaText.Text = $"numbers found: {count}";
        }

        private void CopySum_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(SumResultBox.Text))
                    Clipboard.SetText(SumResultBox.Text);

                StatusText.Text = "sum copied to clipboard.";
            }
            catch (Exception ex)
            {
                StatusText.Text = "copy failed: " + ex.Message;
            }
        }

        private void MainWindow_Closing(object? sender, CancelEventArgs e)
        {
            SettingsStore.Save(Preprocess);
        }
    }
}