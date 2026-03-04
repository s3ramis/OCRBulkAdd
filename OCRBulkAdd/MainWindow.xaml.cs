using System;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using OCRBulkAdd.Processing;
using OCRBulkAdd.Services;
using OCRBulkAdd.Utils;

namespace OCRBulkAdd
{
    public partial class MainWindow : Window
    {
        // preprocessed image as bytes to be sent to tesseract
        private byte[] _currentImageBytes = Array.Empty<byte>();

        // prevent ocr spazzing out for fast concurrent image inputs
        private readonly SemaphoreSlim _ocrLock = new SemaphoreSlim(1, 1);

        private readonly OcrService _ocrService = new OcrService
        {
            // output preprocessed image for debugging
            WriteDebugImages = false
        };

        public MainWindow()
        {
            InitializeComponent();
            UpdateHintVisibility();
        }

        // input handling
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

        private async void RunOcr_Click(object sender, RoutedEventArgs e)
        {
            await RunOcrAsync();
        }

        private void Sum_Click(object sender, RoutedEventArgs e)
        {
            // add numbers extracted from ocr'ed text
            var (sum, count) = NumberExtractor.SumFromText(NumbersTextBox.Text);

            SumResultText.Text = $"sum: {sum.ToString("0.00", CultureInfo.InvariantCulture)} (numbers found: {count})";
        }

        // --- INTERNALS ---
        private void SetImage(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0) return;

            try
            {
                BitmapImage preview = ImageHelper.BytesToBitmapImage(bytes);
                PreviewImage.Source = preview;

                var prepared = ImagePreprocessor.PrepareForOcr(preview);

                _currentImageBytes = ImageHelper.BitmapSourceToPngBytes(prepared);
            }
            catch
            {
                // still try to ocr even if preprocessing fails
                PreviewImage.Source = null;
                _currentImageBytes = bytes;
            }

            OcrPreviewTextBox.Clear();
            NumbersTextBox.Clear();

            StatusText.Text = "image loaded.";
            SumResultText.Text = "sum: -";
            UpdateHintVisibility();
        }

        private void UpdateHintVisibility()
        {
            DropHint.Visibility = PreviewImage.Source == null ? Visibility.Visible : Visibility.Collapsed;
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

                string raw = await _ocrService.RecognizeAsync(_currentImageBytes);

                // preview of ocr'ed text
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
    }
}