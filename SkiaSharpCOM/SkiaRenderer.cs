using SkiaSharp;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
using stdole;
using System.Windows.Forms;
using System.IO;

namespace SkiaSharpCOM
{ 

    [ComVisible(true)]
    public class SkiaRenderer
    {
        private SKBitmap _bitmap;
        public SkiaRenderer() { }
        
        public void CreateBitmap(int width, int height)
        {
            _bitmap = new SKBitmap(width, height);
        }

        [ComVisible(true)]
        public void DrawCircle(float x, float y, float radius, int color, float strokeWidth, bool isAntialias)
        {
            // Create a new SKCanvas object to draw on the bitmap
            using (SKCanvas canvas = new SKCanvas(_bitmap))
            {
                var cor = ColorTranslator.FromOle(color);
                
                // Create an SKPaint object to specify the circle's appearance
                SKPaint paint = new SKPaint
                {
                    Color = new SKColor((uint)cor.ToArgb()),
                    StrokeWidth = strokeWidth,
                    IsAntialias = isAntialias
                };
                canvas.DrawCircle(x, y, radius, paint);
            }
        }

        [ComVisible(true)]
        public IPictureDisp ToPicture()
        {
            SKImage image = SKImage.FromPixels(_bitmap.PeekPixels());

            SKData encoded = image.Encode();
            Stream stream = encoded.AsStream();
            var img = Image.FromStream(stream);
            return (IPictureDisp)Microsoft.VisualBasic.Compatibility.VB6.Support.ImageToIPictureDisp(img);
        }
    }
}