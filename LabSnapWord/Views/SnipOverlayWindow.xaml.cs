using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using LabSnapWord.Services;

namespace LabSnapWord.Views
{
    public partial class SnipOverlayWindow : Window
    {
        private readonly BitmapSource _fullScreenBitmap;
        private bool _isDragging;
        private System.Windows.Point _startPoint;

        public BitmapSource? CroppedResult { get; private set; }

        public SnipOverlayWindow(BitmapSource fullScreenBitmap)
        {
            InitializeComponent();
            _fullScreenBitmap = fullScreenBitmap ?? throw new ArgumentNullException(nameof(fullScreenBitmap));
            BackgroundScreenImage.Source = _fullScreenBitmap;
        }

        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
        }

        private void SnipCanvas_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void SnipCanvas_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _isDragging = true;
            _startPoint = e.GetPosition(SnipCanvas);
            SnipCanvas.CaptureMouse();

            Canvas.SetLeft(SelectionBorder, _startPoint.X);
            Canvas.SetTop(SelectionBorder, _startPoint.Y);
            SelectionBorder.Width = 0;
            SelectionBorder.Height = 0;
            SelectionBorder.Visibility = Visibility.Visible;

            DimensionsBadge.Visibility = Visibility.Visible;
            UpdateSelectionVisual(_startPoint);
        }

        private void SnipCanvas_MouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (!_isDragging)
            {
                return;
            }

            System.Windows.Point currentPoint = e.GetPosition(SnipCanvas);
            UpdateSelectionVisual(currentPoint);
        }

        private void SnipCanvas_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isDragging)
            {
                return;
            }

            _isDragging = false;
            SnipCanvas.ReleaseMouseCapture();

            System.Windows.Point endPoint = e.GetPosition(SnipCanvas);
            Int32Rect pixelRect = ConvertCanvasRectToPixelRect(_startPoint, endPoint);

            if (pixelRect.Width < 8 || pixelRect.Height < 8)
            {
                SelectionBorder.Visibility = Visibility.Collapsed;
                DimensionsBadge.Visibility = Visibility.Collapsed;
                return;
            }

            CroppedResult = ScreenCaptureService.CropBitmap(_fullScreenBitmap, pixelRect);
            DialogResult = true;
            Close();
        }

        private void UpdateSelectionVisual(System.Windows.Point currentPoint)
        {
            double x = Math.Min(_startPoint.X, currentPoint.X);
            double y = Math.Min(_startPoint.Y, currentPoint.Y);
            double width = Math.Abs(currentPoint.X - _startPoint.X);
            double height = Math.Abs(currentPoint.Y - _startPoint.Y);

            Canvas.SetLeft(SelectionBorder, x);
            Canvas.SetTop(SelectionBorder, y);
            SelectionBorder.Width = width;
            SelectionBorder.Height = height;

            Int32Rect pixelRect = ConvertCanvasRectToPixelRect(_startPoint, currentPoint);
            DimensionsText.Text = $"{pixelRect.Width} × {pixelRect.Height} px";

            double badgeTop = y > 36 ? y - 32 : y + height + 8;
            Canvas.SetLeft(DimensionsBadge, x);
            Canvas.SetTop(DimensionsBadge, badgeTop);
        }

        private Int32Rect ConvertCanvasRectToPixelRect(System.Windows.Point p1, System.Windows.Point p2)
        {
            double canvasW = Math.Max(1.0, SnipCanvas.ActualWidth);
            double canvasH = Math.Max(1.0, SnipCanvas.ActualHeight);

            double scaleX = _fullScreenBitmap.PixelWidth / canvasW;
            double scaleY = _fullScreenBitmap.PixelHeight / canvasH;

            double left = Math.Max(0, Math.Min(p1.X, p2.X));
            double top = Math.Max(0, Math.Min(p1.Y, p2.Y));
            double width = Math.Abs(p2.X - p1.X);
            double height = Math.Abs(p2.Y - p1.Y);

            int px = (int)Math.Round(left * scaleX);
            int py = (int)Math.Round(top * scaleY);
            int pw = (int)Math.Round(width * scaleX);
            int ph = (int)Math.Round(height * scaleY);

            return new Int32Rect(px, py, pw, ph);
        }
    }
}
