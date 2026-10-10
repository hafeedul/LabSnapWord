using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media.Imaging;

namespace LabSnapWord.Services
{
    public static class ScreenCaptureService
    {
        /// <summary>
        /// Captures the entire primary screen as a WPF BitmapSource.
        /// </summary>
        public static BitmapSource CaptureFullScreen(out Rectangle screenBounds)
        {
            Screen primary = Screen.PrimaryScreen ?? Screen.AllScreens[0];
            screenBounds = primary.Bounds;

            using (Bitmap bmp = new Bitmap(screenBounds.Width, screenBounds.Height, PixelFormat.Format32bppArgb))
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen(
                        screenBounds.Left,
                        screenBounds.Top,
                        0,
                        0,
                        screenBounds.Size,
                        CopyPixelOperation.SourceCopy);
                }

                return ConvertDrawingBitmapToBitmapSource(bmp);
            }
        }

        /// <summary>
        /// Crops a WPF BitmapSource to the specified pixel rectangle, clamping to valid bounds.
        /// </summary>
        public static BitmapSource CropBitmap(BitmapSource source, Int32Rect rect)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            int x = Math.Max(0, Math.Min(rect.X, source.PixelWidth - 1));
            int y = Math.Max(0, Math.Min(rect.Y, source.PixelHeight - 1));
            int width = Math.Max(1, Math.Min(rect.Width, source.PixelWidth - x));
            int height = Math.Max(1, Math.Min(rect.Height, source.PixelHeight - y));

            CroppedBitmap cropped = new CroppedBitmap(source, new Int32Rect(x, y, width, height));
            cropped.Freeze();
            return cropped;
        }

        /// <summary>
        /// Encodes a WPF BitmapSource into PNG byte array for Word document insertion.
        /// </summary>
        public static byte[] EncodeToPngBytes(BitmapSource source)
        {
            PngBitmapEncoder encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(source));

            using (MemoryStream ms = new MemoryStream())
            {
                encoder.Save(ms);
                return ms.ToArray();
            }
        }

        private static BitmapSource ConvertDrawingBitmapToBitmapSource(Bitmap bitmap)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                bitmap.Save(ms, ImageFormat.Png);
                ms.Position = 0;

                BitmapImage bmpImage = new BitmapImage();
                bmpImage.BeginInit();
                bmpImage.CacheOption = BitmapCacheOption.OnLoad;
                bmpImage.StreamSource = ms;
                bmpImage.EndInit();
                bmpImage.Freeze();
                return bmpImage;
            }
        }
    }
}
