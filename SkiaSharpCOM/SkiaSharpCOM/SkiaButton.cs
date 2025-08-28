using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SkiaSharp;
using stdole;

namespace SkiaSharpCOM
{


    [ComVisible(true)]
    [Guid("B8B3E4D2-1234-4567-8901-123456789ABC")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ISkiaButton
    {
        // Propriedades do botão
        string Text { get; set; }
        int Width { get; set; }
        int Height { get; set; }
        int BackgroundColor { get; set; }
        int TextColor { get; set; }
        int BorderColor { get; set; }
        float BorderWidth { get; set; }
        float CornerRadius { get; set; }
        string FontFamily { get; set; }
        float FontSize { get; set; }
        bool Bold { get; set; }
        bool IsPressed { get; set; }
        bool IsHovered { get; set; }
        bool Enabled { get; set; }

        // Métodos
        [ComVisible(true)]
        stdole.IPictureDisp RenderButton();
        [ComVisible(true)]
        stdole.IPictureDisp RenderButtonState(bool pressed, bool hovered, bool enabled);
        [ComVisible(true)] 
        void SetGradientBackground(int startColor, int endColor);

        [ComVisible(true)] 
        void SetTextShadow(float offsetX, float offsetY, float blurRadius, int shadowColor);
    }

    [ComVisible(true)]
    [Guid("C9C4F5E3-2345-5678-9012-234567890BCD")]
    [ClassInterface(ClassInterfaceType.None)]
    public class SkiaButton : ISkiaButton
    {
        private string _text = "Button";
        private int _width = 120;
        private int _height = 40;
        private uint _backgroundColor = 0xFF4285F4; // Azul Google
        private uint _textColor = 0xFFFFFFFF; // Branco
        private uint _borderColor = 0xFF1976D2; // Azul mais escuro
        private float _borderWidth = 2.0f;
        private float _cornerRadius = 8.0f;
        private string _fontFamily = "Segoe UI";
        private float _fontSize = 14.0f;
        private bool _bold = false;
        private bool _isPressed = false;
        private bool _isHovered = false;
        private bool _enabled = true;

        // Propriedades para gradiente
        private bool _useGradient = false;
        private uint _gradientStartColor = 0xFF4285F4;
        private uint _gradientEndColor = 0xFF1976D2;

        // Propriedades para sombra do texto
        private bool _useTextShadow = false;
        private float _shadowOffsetX = 1.0f;
        private float _shadowOffsetY = 1.0f;
        private float _shadowBlurRadius = 2.0f;
        private uint _shadowColor = 0x80000000; // Preto com 50% transparência

        #region Propriedades COM

        public string Text
        {
            get => _text;
            set => _text = value ?? "Button";
        }

        public int Width
        {
            get => _width;
            set => _width = Math.Max(1, value);
        }

        public int Height
        {
            get => _height;
            set => _height = Math.Max(1, value);
        }

        public int BackgroundColor
        {
            get => unchecked((int) _backgroundColor);
            set => _backgroundColor = unchecked((uint)value);  
        }

        [ComVisible(true)]
        public int TextColor
        {
            get => unchecked((int)_textColor);
            set => _textColor =  unchecked((uint)value);  
        }

        public int BorderColor
        {
            get => unchecked((int) _borderColor);
            set => _borderColor =  unchecked((uint)value);  
        }

        public float BorderWidth
        {
            get => _borderWidth;
            set => _borderWidth = Math.Max(0, value);
        }

        public float CornerRadius
        {
            get => _cornerRadius;
            set => _cornerRadius = Math.Max(0, value);
        }

        public string FontFamily
        {
            get => _fontFamily;
            set => _fontFamily = value ?? "Segoe UI";
        }

        public float FontSize
        {
            get => _fontSize;
            set => _fontSize = Math.Max(1, value);
        }

        public bool Bold
        {
            get => _bold;
            set => _bold = value;
        }

        public bool IsPressed
        {
            get => _isPressed;
            set => _isPressed = value;
        }

        public bool IsHovered
        {
            get => _isHovered;
            set => _isHovered = value;
        }

        public bool Enabled
        {
            get => _enabled;
            set => _enabled = value;
        }

        #endregion

        #region Métodos COM

        [ComVisible(true)]
        public stdole.IPictureDisp RenderButton()
        {
            return RenderButtonState(_isPressed, _isHovered, _enabled);
        }

        [ComVisible(true)]
        public IPictureDisp RenderButtonState(bool pressed, bool hovered, bool enabled)
        {
            SKBitmap bitmap = new SKBitmap(_width, _height);
            SKCanvas canvas = new SKCanvas(bitmap);

           // var surface = SKSurface.Create(new SKImageInfo(_width, _height));
           // var canvas = surface.Canvas;

            canvas.Clear(SKColors.Transparent);

            // Calcular cores baseadas no estado
            var bgColor = CalculateBackgroundColor(pressed, hovered, enabled);
            var txtColor = CalculateTextColor(enabled);
            var brdColor = CalculateBorderColor(pressed, hovered, enabled);

            // Desenhar o fundo do botão
            DrawBackground(canvas, bgColor, pressed);

            // Desenhar borda se especificada
            if (_borderWidth > 0)
            {
                DrawBorder(canvas, brdColor);
            }

            // Desenhar texto
            DrawText(canvas, txtColor, pressed);
           
            // Converter para IPictureDisp
            return ToPicture(bitmap);
                /*
            using (var data = image.Encode(SKEncodedImageFormat.Png, 100))
            {
                var bytes = data.ToArray();
                return ConvertToPicture(bytes);
            }*/
        }
        public IPictureDisp ToPicture(SKBitmap bitmap)
        {
            SKImage image = SKImage.FromPixels(bitmap.PeekPixels());

            SKData encoded = image.Encode();
            Stream stream = encoded.AsStream();
            var img = Image.FromStream(stream);
            return ImageToIPictureConverter.ConvertImageToIPictureDisp(img); ;
        }
        [ComVisible(true)]
        public void SetGradientBackground(int startColor, int endColor)
        {
            _useGradient = true;
            _gradientStartColor = unchecked((uint)startColor);  
            _gradientEndColor = unchecked((uint)endColor);
        }

        [ComVisible(true)]
        public void SetTextShadow(float offsetX, float offsetY, float blurRadius, int shadowColor)
        {
            _useTextShadow = true;
            _shadowOffsetX = offsetX;
            _shadowOffsetY = offsetY;
            _shadowBlurRadius = blurRadius;
            _shadowColor = unchecked((uint)shadowColor); 
        }

        #endregion

        #region Métodos de Renderização

        private void DrawBackground(SKCanvas canvas, uint color, bool pressed)
        {
            var rect = new SKRect(0, 0, _width, _height);

            if (_cornerRadius > 0)
            {
                rect.Inflate(-_borderWidth / 2, -_borderWidth / 2);
            }

            using (var paint = new SKPaint())
            {
                paint.IsAntialias = true;

                if (_useGradient)
                {
                    // Criar gradiente vertical
                    var startColor = pressed ? ModifyColor(_gradientEndColor, 0.9f) : new SKColor(_gradientStartColor);
                    var endColor = pressed ? ModifyColor(_gradientStartColor, 0.9f) : new SKColor(_gradientEndColor);

                    paint.Shader = SKShader.CreateLinearGradient(
                        new SKPoint(rect.Left, rect.Top),
                        new SKPoint(rect.Left, rect.Bottom),
                        new SKColor[] { startColor, endColor },
                        null,
                        SKShaderTileMode.Clamp);
                }
                else
                {
                    paint.Color = new SKColor(color);
                }

                // Adicionar efeito de profundidade se pressionado
                if (pressed)
                {
                    paint.ColorFilter = SKColorFilter.CreateColorMatrix(new float[]
                    {
                        0.8f, 0, 0, 0, 0,
                        0, 0.8f, 0, 0, 0,
                        0, 0, 0.8f, 0, 0,
                        0, 0, 0, 1, 0
                    });
                }

                if (_cornerRadius > 0)
                {
                    canvas.DrawRoundRect(rect, _cornerRadius, _cornerRadius, paint);
                }
                else
                {
                    canvas.DrawRect(rect, paint);
                }
            }
        }

        private void DrawBorder(SKCanvas canvas, uint color)
        {
            var rect = new SKRect(_borderWidth / 2, _borderWidth / 2,
                                _width - _borderWidth / 2, _height - _borderWidth / 2);

            using (var paint = new SKPaint())
            {
                paint.IsAntialias = true;
                paint.Style = SKPaintStyle.Stroke;
                paint.Color = new SKColor(color);
                paint.StrokeWidth = _borderWidth;

                if (_cornerRadius > 0)
                {
                    canvas.DrawRoundRect(rect, _cornerRadius, _cornerRadius, paint);
                }
                else
                {
                    canvas.DrawRect(rect, paint);
                }
            }
        }

        private void DrawText(SKCanvas canvas, uint color, bool pressed)
        {
            if (string.IsNullOrEmpty(_text))
                return;

            using (var paint = new SKPaint())
            {
                paint.IsAntialias = true;
                paint.Color = new SKColor(color);
                paint.TextSize = _fontSize;
                paint.Typeface = SKTypeface.FromFamilyName(_fontFamily,
                    _bold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal,
                    SKFontStyleWidth.Normal,
                    SKFontStyleSlant.Upright);

                // Desenhar sombra do texto se especificada
                if (_useTextShadow)
                {
                    using (var shadowPaint = paint.Clone())
                    {
                        shadowPaint.Color = new SKColor(_shadowColor);
                        if (_shadowBlurRadius > 0)
                        {
                            shadowPaint.MaskFilter = SKMaskFilter.CreateBlur(SKBlurStyle.Normal, _shadowBlurRadius);
                        }

                        var shadowBounds = new SKRect();
                        shadowPaint.MeasureText(_text, ref shadowBounds);

                        var shadowX = (_width - shadowBounds.Width) / 2 + _shadowOffsetX;
                        var shadowY = (_height - shadowBounds.Height) / 2 - shadowBounds.Top + _shadowOffsetY;

                        canvas.DrawText(_text, shadowX, shadowY, shadowPaint);
                    }
                }

                // Calcular posição centralizada do texto
                var textBounds = new SKRect();
                paint.MeasureText(_text, ref textBounds);

                var textX = (_width - textBounds.Width) / 2;
                var textY = (_height - textBounds.Height) / 2 - textBounds.Top;

                // Ajustar posição se pressionado para efeito 3D
                if (pressed)
                {
                    textX += 1;
                    textY += 1;
                }

                canvas.DrawText(_text, textX, textY, paint);
            }
        }

        #endregion

        #region Métodos Auxiliares

        private uint CalculateBackgroundColor(bool pressed, bool hovered, bool enabled)
        { 
            
            if (!enabled)
                return (uint) ModifyColor(_backgroundColor, 0.5f);

            if (pressed)
                return (uint) ModifyColor(_backgroundColor, 0.8f);

            if (hovered)
                return (uint) ModifyColor(_backgroundColor, 1.1f);

            return _backgroundColor;
        }

        private uint CalculateTextColor(bool enabled)
        {
            if (!enabled)
                return (uint) ModifyColor(_textColor, 0.6f);

            return _textColor;
        }

        private uint CalculateBorderColor(bool pressed, bool hovered, bool enabled)
        {

            if (!enabled)
                return (uint)ModifyColor(_borderColor, 0.5f);

            if (pressed)
                return (uint) ModifyColor(_borderColor, 0.8f);

            if (hovered)
                return (uint) ModifyColor(_borderColor, 1.1f);

            return _borderColor;
        }

        private SKColor ModifyColor(uint color, float factor)
        {
            var skColor = new SKColor(color);
            var r = (byte)Math.Min(255, (int)(skColor.Red * factor));
            var g = (byte)Math.Min(255, (int)(skColor.Green * factor));
            var b = (byte)Math.Min(255, (int)(skColor.Blue * factor));

            return new SKColor(r, g, b, skColor.Alpha);
        }
        [ComVisible(true)]
        public stdole.IPictureDisp ToPicture(SKImage image)
        {
            SKData encoded = image.Encode();
            Stream stream = encoded.AsStream();
            var img = Image.FromStream(stream);
            return  ImageToIPictureConverter.ConvertImageToIPictureDisp(img);
        }
        #endregion
    }
   
}