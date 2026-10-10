using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using LabSnapWord.Services;
using LabSnapWord.Views;

namespace LabSnapWord
{
    public partial class MainWindow : Window
    {
        // Win32 Global Hotkey Constants (Ctrl + Shift + S)
        private const int HOTKEY_ID_SNIP = 9001;
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;
        private const uint MOD_NOREPEAT = 0x4000;
        private const uint VK_S = 0x53;
        private const int WM_HOTKEY = 0x0312;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private IntPtr _windowHandle;
        private HwndSource? _hwndSource;
        private bool _isSnippingInProgress;

        private string? _selectedWordFilePath;
        private BitmapSource? _originalImage;
        private BitmapSource? _currentImage;

        private bool _isDraggingCrop;
        private System.Windows.Point _cropStartPoint;
        private Int32Rect? _activeCropRect;

        private int _currentTaskNumber = 1;
        private int _sessionAddedCount = 0;

        public MainWindow()
        {
            InitializeComponent();

            if (AppInstallerService.IsRunningFromInstalledLocation())
            {
                InstallAppButton.Content = "✓ Recreate Shortcuts";
            }
        }

        #region Global Hotkey (Ctrl + Shift + S)

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);

            _windowHandle = new WindowInteropHelper(this).Handle;
            _hwndSource = HwndSource.FromHwnd(_windowHandle);
            _hwndSource?.AddHook(HwndHook);

            bool registered = RegisterHotKey(
                _windowHandle,
                HOTKEY_ID_SNIP,
                MOD_CONTROL | MOD_SHIFT | MOD_NOREPEAT,
                VK_S);

            if (registered)
            {
                SetStatus("Ready. Press Ctrl + Shift + S anytime (even in the background) to snip a screenshot!", isSuccess: false);
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            if (_windowHandle != IntPtr.Zero)
            {
                UnregisterHotKey(_windowHandle, HOTKEY_ID_SNIP);
            }
            _hwndSource?.RemoveHook(HwndHook);
            base.OnClosed(e);
        }

        private IntPtr HwndHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_HOTKEY && wParam.ToInt32() == HOTKEY_ID_SNIP)
            {
                handled = true;
                if (!_isSnippingInProgress)
                {
                    _ = StartScreenRegionSnipAsync();
                }
            }
            return IntPtr.Zero;
        }

        private void Window_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.S &&
                (Keyboard.Modifiers & (ModifierKeys.Control | ModifierKeys.Shift)) == (ModifierKeys.Control | ModifierKeys.Shift))
            {
                e.Handled = true;
                if (!_isSnippingInProgress)
                {
                    _ = StartScreenRegionSnipAsync();
                }
            }
        }

        #endregion

        #region Header & Word File Selection

        private void InstallAppButton_Click(object sender, RoutedEventArgs e)
        {
            InstallerWindow installer = new InstallerWindow(isStandaloneSetupMode: false)
            {
                Owner = this
            };
            installer.ShowDialog();
        }

        private void AlwaysOnTopCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            Topmost = AlwaysOnTopCheckBox.IsChecked == true;
        }

        private void ChooseWordFileButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select Existing Word Report (.docx)",
                Filter = "Word Document (*.docx)|*.docx",
                CheckFileExists = true
            };

            if (dialog.ShowDialog(this) == true)
            {
                SetActiveWordFile(dialog.FileName);
                SetStatus($"Selected Word file: {System.IO.Path.GetFileName(dialog.FileName)} (works even while open in Word!)", isSuccess: true);
            }
        }

        private void CreateWordFileButton_Click(object sender, RoutedEventArgs e)
        {
            PromptCreateOrSelectWordFile();
        }

        private bool PromptCreateOrSelectWordFile()
        {
            Microsoft.Win32.SaveFileDialog dialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Create or Choose Word Report (.docx)",
                Filter = "Word Document (*.docx)|*.docx",
                FileName = "Lab_Report.docx",
                DefaultExt = ".docx",
                OverwritePrompt = true
            };

            if (dialog.ShowDialog(this) == true)
            {
                try
                {
                    WordReportService.EnsureOrCreateDocument(dialog.FileName, reportTitle: null, overwrite: true);
                    SetActiveWordFile(dialog.FileName);
                    SetStatus($"Ready to write to: {System.IO.Path.GetFileName(dialog.FileName)}", isSuccess: true);
                    return true;
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show(
                        this,
                        $"Could not create the Word file:\n{ex.Message}",
                        "File Error",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }

            return false;
        }

        private void SetActiveWordFile(string filePath)
        {
            _selectedWordFilePath = filePath;
            WordFilePathTextBox.Text = filePath;
            WordFilePathTextBox.ToolTip = filePath;
            WordFilePathTextBox.Foreground = (SolidColorBrush)FindResource("TextPrimaryBrush");
            OpenInWordButton.IsEnabled = true;
        }

        private void OpenInWordButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_selectedWordFilePath) || !File.Exists(_selectedWordFilePath))
            {
                System.Windows.MessageBox.Show(
                    this,
                    "Please select or create a Word (.docx) file first.",
                    "No Word File",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo(_selectedWordFilePath)
                {
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    this,
                    $"Could not open file in Microsoft Word:\n{ex.Message}",
                    "Open Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        #endregion

        #region Screenshot Capture Actions

        private async void SnipRegionButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_isSnippingInProgress)
            {
                await StartScreenRegionSnipAsync();
            }
        }

        private async Task StartScreenRegionSnipAsync()
        {
            if (_isSnippingInProgress)
            {
                return;
            }

            _isSnippingInProgress = true;
            WindowState previousState = WindowState == WindowState.Minimized ? WindowState.Normal : WindowState;
            bool wasTopmost = Topmost;

            try
            {
                Hide();
                await Task.Delay(240);

                BitmapSource fullScreen = ScreenCaptureService.CaptureFullScreen(out _);
                SnipOverlayWindow snipWindow = new SnipOverlayWindow(fullScreen);
                bool? result = snipWindow.ShowDialog();

                Show();
                WindowState = previousState;
                Topmost = wasTopmost;
                Activate();

                if (result == true && snipWindow.CroppedResult != null)
                {
                    LoadCapturedBitmap(fullScreen, snipWindow.CroppedResult);
                    SetStatus("Region captured! Drag on the preview to crop further, or click 'Add Screenshot to Word'.", isSuccess: true);
                }
                else
                {
                    SetStatus("Screen snip cancelled. Press Ctrl + Shift + S anytime to snip again.", isSuccess: false);
                }
            }
            catch (Exception ex)
            {
                Show();
                WindowState = previousState;
                Topmost = wasTopmost;
                System.Windows.MessageBox.Show(
                    this,
                    $"Failed to capture screen region:\n{ex.Message}",
                    "Capture Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                _isSnippingInProgress = false;
            }
        }

        private async void CaptureFullScreenButton_Click(object sender, RoutedEventArgs e)
        {
            WindowState previousState = WindowState;
            bool wasTopmost = Topmost;

            try
            {
                Hide();
                await Task.Delay(240);

                BitmapSource fullScreen = ScreenCaptureService.CaptureFullScreen(out _);

                Show();
                WindowState = previousState;
                Topmost = wasTopmost;
                Activate();

                LoadCapturedBitmap(fullScreen, fullScreen);
                SetStatus("Full screen captured! Drag a box on the image to crop, or click 'Add Screenshot to Word'.", isSuccess: true);
            }
            catch (Exception ex)
            {
                Show();
                WindowState = previousState;
                Topmost = wasTopmost;
                System.Windows.MessageBox.Show(
                    this,
                    $"Failed to capture full screen:\n{ex.Message}",
                    "Capture Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void PasteClipboardButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!System.Windows.Clipboard.ContainsImage())
                {
                    System.Windows.MessageBox.Show(
                        this,
                        "No image found on the Windows Clipboard.\n\nTip: Press Ctrl + Shift + S to snip directly, or press PrintScreen and click 'Clipboard'.",
                        "Clipboard Empty",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }

                BitmapSource? clipboardImage = System.Windows.Clipboard.GetImage();
                if (clipboardImage != null)
                {
                    clipboardImage.Freeze();
                    LoadCapturedBitmap(clipboardImage, clipboardImage);
                    SetStatus("Image pasted from clipboard! Drag on the preview to crop if needed.", isSuccess: true);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    this,
                    $"Could not read image from clipboard:\n{ex.Message}",
                    "Clipboard Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
        }

        private void LoadImageFileButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Open Screenshot Image",
                Filter = "Image Files (*.png;*.jpg;*.jpeg;*.bmp)|*.png;*.jpg;*.jpeg;*.bmp"
            };

            if (dialog.ShowDialog(this) == true)
            {
                try
                {
                    BitmapImage bmp = new BitmapImage();
                    bmp.BeginInit();
                    bmp.CacheOption = BitmapCacheOption.OnLoad;
                    bmp.UriSource = new Uri(dialog.FileName);
                    bmp.EndInit();
                    bmp.Freeze();

                    LoadCapturedBitmap(bmp, bmp);
                    SetStatus($"Loaded image: {System.IO.Path.GetFileName(dialog.FileName)}", isSuccess: true);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show(
                        this,
                        $"Could not load image file:\n{ex.Message}",
                        "Image Error",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }
        }

        private void LoadCapturedBitmap(BitmapSource originalBitmap, BitmapSource currentBitmap)
        {
            _originalImage = originalBitmap;
            SetCurrentPreviewBitmap(currentBitmap);
        }

        private void SetCurrentPreviewBitmap(BitmapSource bitmap)
        {
            _currentImage = bitmap;
            ClearActiveCropSelection();

            EmptyStatePanel.Visibility = Visibility.Collapsed;
            PreviewViewbox.Visibility = Visibility.Visible;

            PreviewImage.Source = _currentImage;
            ImageWorkspaceGrid.Width = _currentImage.PixelWidth;
            ImageWorkspaceGrid.Height = _currentImage.PixelHeight;
            CropOverlayCanvas.Width = _currentImage.PixelWidth;
            CropOverlayCanvas.Height = _currentImage.PixelHeight;

            double dynamicThickness = Math.Max(2.5, Math.Max(_currentImage.PixelWidth, _currentImage.PixelHeight) / 340.0);
            CropSelectionBorder.BorderThickness = new Thickness(dynamicThickness);

            AddToWordButton.IsEnabled = true;
            ResetCropButton.IsEnabled = _originalImage != null && !ReferenceEquals(_originalImage, _currentImage);
            UpdateDimensionsLabel();
        }

        #endregion

        #region Interactive In-App Crop Canvas with Dark Outside Masks

        private void CropOverlayCanvas_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (_currentImage == null)
            {
                return;
            }

            _isDraggingCrop = true;
            _cropStartPoint = ClampPointToImage(e.GetPosition(CropOverlayCanvas));
            CropOverlayCanvas.CaptureMouse();

            CropSelectionBorder.Visibility = Visibility.Visible;
            MaskTop.Visibility = Visibility.Visible;
            MaskBottom.Visibility = Visibility.Visible;
            MaskLeft.Visibility = Visibility.Visible;
            MaskRight.Visibility = Visibility.Visible;

            UpdateInAppCropBox(_cropStartPoint, _cropStartPoint);
        }

        private void CropOverlayCanvas_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (!_isDraggingCrop || _currentImage == null)
            {
                return;
            }

            System.Windows.Point currentPoint = ClampPointToImage(e.GetPosition(CropOverlayCanvas));
            UpdateInAppCropBox(_cropStartPoint, currentPoint);
        }

        private void CropOverlayCanvas_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isDraggingCrop || _currentImage == null)
            {
                return;
            }

            _isDraggingCrop = false;
            CropOverlayCanvas.ReleaseMouseCapture();

            System.Windows.Point endPoint = ClampPointToImage(e.GetPosition(CropOverlayCanvas));
            UpdateInAppCropBox(_cropStartPoint, endPoint);

            if (_activeCropRect == null || _activeCropRect.Value.Width < 10 || _activeCropRect.Value.Height < 10)
            {
                ClearActiveCropSelection();
            }
            else
            {
                ApplyCropButton.IsEnabled = true;
                SetStatus($"Crop area selected ({_activeCropRect.Value.Width} × {_activeCropRect.Value.Height} px). Click 'Apply Crop' or 'Add Screenshot to Word'.", isSuccess: false);
            }
        }

        private void CropOverlayCanvas_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            ClearActiveCropSelection();
        }

        private void UpdateInAppCropBox(System.Windows.Point p1, System.Windows.Point p2)
        {
            if (_currentImage == null)
            {
                return;
            }

            double imgW = _currentImage.PixelWidth;
            double imgH = _currentImage.PixelHeight;

            double left = Math.Min(p1.X, p2.X);
            double top = Math.Min(p1.Y, p2.Y);
            double width = Math.Abs(p2.X - p1.X);
            double height = Math.Abs(p2.Y - p1.Y);
            double right = left + width;
            double bottom = top + height;

            Canvas.SetLeft(CropSelectionBorder, left);
            Canvas.SetTop(CropSelectionBorder, top);
            CropSelectionBorder.Width = width;
            CropSelectionBorder.Height = height;

            SetRectBounds(MaskTop, 0, 0, imgW, top);
            SetRectBounds(MaskBottom, 0, bottom, imgW, Math.Max(0, imgH - bottom));
            SetRectBounds(MaskLeft, 0, top, left, height);
            SetRectBounds(MaskRight, right, top, Math.Max(0, imgW - right), height);

            int px = (int)Math.Round(left);
            int py = (int)Math.Round(top);
            int pw = (int)Math.Round(width);
            int ph = (int)Math.Round(height);

            _activeCropRect = new Int32Rect(px, py, pw, ph);
            UpdateDimensionsLabel();
        }

        private static void SetRectBounds(Rectangle rect, double left, double top, double width, double height)
        {
            Canvas.SetLeft(rect, left);
            Canvas.SetTop(rect, top);
            rect.Width = Math.Max(0, width);
            rect.Height = Math.Max(0, height);
        }

        private System.Windows.Point ClampPointToImage(System.Windows.Point pt)
        {
            if (_currentImage == null)
            {
                return new System.Windows.Point(0, 0);
            }

            double x = Math.Max(0, Math.Min(pt.X, _currentImage.PixelWidth));
            double y = Math.Max(0, Math.Min(pt.Y, _currentImage.PixelHeight));
            return new System.Windows.Point(x, y);
        }

        private void ClearActiveCropSelection()
        {
            _isDraggingCrop = false;
            _activeCropRect = null;
            CropSelectionBorder.Visibility = Visibility.Collapsed;
            MaskTop.Visibility = Visibility.Collapsed;
            MaskBottom.Visibility = Visibility.Collapsed;
            MaskLeft.Visibility = Visibility.Collapsed;
            MaskRight.Visibility = Visibility.Collapsed;
            ApplyCropButton.IsEnabled = false;
            UpdateDimensionsLabel();
        }

        private void ApplyCropButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentImage == null || _activeCropRect == null)
            {
                return;
            }

            BitmapSource cropped = ScreenCaptureService.CropBitmap(_currentImage, _activeCropRect.Value);
            SetCurrentPreviewBitmap(cropped);
            ResetCropButton.IsEnabled = true;
            SetStatus($"Cropped image to {cropped.PixelWidth} × {cropped.PixelHeight} px.", isSuccess: true);
        }

        private void ResetCropButton_Click(object sender, RoutedEventArgs e)
        {
            if (_originalImage == null)
            {
                return;
            }

            SetCurrentPreviewBitmap(_originalImage);
            SetStatus("Restored original un-cropped screenshot.", isSuccess: false);
        }

        private void UpdateDimensionsLabel()
        {
            if (_currentImage == null)
            {
                ImageDimensionsText.Text = "Press Ctrl + Shift + S anytime to snip a screenshot, then drag here to crop";
                return;
            }

            if (_activeCropRect != null && _activeCropRect.Value.Width >= 10 && _activeCropRect.Value.Height >= 10)
            {
                ImageDimensionsText.Text = $"Image: {_currentImage.PixelWidth} × {_currentImage.PixelHeight} px   •   Selected Crop: {_activeCropRect.Value.Width} × {_activeCropRect.Value.Height} px";
            }
            else
            {
                ImageDimensionsText.Text = $"Image Size: {_currentImage.PixelWidth} × {_currentImage.PixelHeight} px   •   Drag anywhere on image to crop";
            }
        }

        #endregion

        #region Task Numbering & Add to Word (Supports Open Word Files!)

        private void NextTaskNumberButton_Click(object sender, RoutedEventArgs e)
        {
            AdvanceTaskNumberInTextBox();
        }

        private void AdvanceTaskNumberInTextBox()
        {
            string current = TaskHeadingTextBox.Text ?? string.Empty;
            Match match = Regex.Match(current, @"^(.*?\b(?:Task|Step|Question|Q|Lab)\s*)(\d+)(.*)$", RegexOptions.IgnoreCase);
            if (match.Success && int.TryParse(match.Groups[2].Value, out int parsedNum))
            {
                _currentTaskNumber = parsedNum + 1;
                TaskHeadingTextBox.Text = $"{match.Groups[1].Value}{_currentTaskNumber}: ";
            }
            else
            {
                _currentTaskNumber++;
                TaskHeadingTextBox.Text = $"Task {_currentTaskNumber}: ";
            }

            TaskHeadingTextBox.CaretIndex = TaskHeadingTextBox.Text.Length;
        }

        private void AddToWordButton_Click(object sender, RoutedEventArgs e)
        {
            if (_currentImage == null)
            {
                System.Windows.MessageBox.Show(
                    this,
                    "Please capture or load a screenshot first (or press Ctrl + Shift + S).",
                    "No Screenshot",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            if (_activeCropRect != null && _activeCropRect.Value.Width >= 10 && _activeCropRect.Value.Height >= 10)
            {
                BitmapSource cropped = ScreenCaptureService.CropBitmap(_currentImage, _activeCropRect.Value);
                SetCurrentPreviewBitmap(cropped);
                ResetCropButton.IsEnabled = true;
            }

            if (string.IsNullOrWhiteSpace(_selectedWordFilePath))
            {
                bool chosen = PromptCreateOrSelectWordFile();
                if (!chosen || string.IsNullOrWhiteSpace(_selectedWordFilePath))
                {
                    return;
                }
            }

            try
            {
                AddToWordButton.IsEnabled = false;

                byte[] pngBytes = ScreenCaptureService.EncodeToPngBytes(_currentImage);
                string? heading = string.IsNullOrWhiteSpace(TaskHeadingTextBox.Text) ? null : TaskHeadingTextBox.Text.Trim();
                string? notes = string.IsNullOrWhiteSpace(TaskNotesTextBox.Text) ? null : TaskNotesTextBox.Text.Trim();

                bool insertedLiveInWord = WordReportService.AppendScreenshotEntry(
                    _selectedWordFilePath,
                    pngBytes,
                    _currentImage.PixelWidth,
                    _currentImage.PixelHeight,
                    heading,
                    notes);

                _sessionAddedCount++;
                SessionCounterText.Text = $"Total added this session: {_sessionAddedCount}";

                string fileName = System.IO.Path.GetFileName(_selectedWordFilePath);
                string headingDisplay = heading ?? "Screenshot";
                string modeLabel = insertedLiveInWord ? "live into open Word window" : $"to {fileName}";

                SetStatus($"✅ Added \"{headingDisplay}\" ({_currentImage.PixelWidth}×{_currentImage.PixelHeight} px) {modeLabel}!", isSuccess: true);

                string historyEntry = $"✓ {DateTime.Now:HH:mm} — {headingDisplay} ({_currentImage.PixelWidth}×{_currentImage.PixelHeight}px)";
                SessionHistoryListBox.Items.Insert(0, historyEntry);

                TaskNotesTextBox.Text = string.Empty;
                if (AutoIncrementTaskCheckBox.IsChecked == true && !string.IsNullOrWhiteSpace(heading))
                {
                    AdvanceTaskNumberInTextBox();
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(
                    this,
                    $"Could not add screenshot to Word:\n{ex.Message}",
                    "Word Insert Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
            }
            finally
            {
                AddToWordButton.IsEnabled = _currentImage != null;
            }
        }

        private void SetStatus(string message, bool isSuccess)
        {
            StatusBannerText.Text = message;
            StatusBannerText.Foreground = isSuccess
                ? new SolidColorBrush(System.Windows.Media.Color.FromRgb(16, 185, 129))
                : (SolidColorBrush)FindResource("TextSecondaryBrush");
        }

        #endregion
    }
}
