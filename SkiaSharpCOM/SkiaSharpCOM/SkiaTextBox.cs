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
    [Guid("D9D4E5F3-3456-6789-0123-345678901CDE")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ISkiaTextBox
    {
        // Propriedades do texto
        string Text { get; set; }
        string PlaceholderText { get; set; }
        int MaxLength { get; set; }
        bool ReadOnly { get; set; }
        bool Multiline { get; set; }
        int TextAlign { get; set; } // 0=Left, 1=Center, 2=Right

        // Propriedades de aparência
        int Width { get; set; }
        int Height { get; set; }
        int BackgroundColor { get; set; }
        int TextColor { get; set; }
        int PlaceholderColor { get; set; }
        int BorderColor { get; set; }
        int FocusBorderColor { get; set; }
        int SelectionColor { get; set; }
        float BorderWidth { get; set; }
        float CornerRadius { get; set; }

        // Propriedades da fonte
        string FontFamily { get; set; }
        float FontSize { get; set; }
        bool Bold { get; set; }

        // Estado
        bool HasFocus { get; set; }
        bool Enabled { get; set; }
        int CursorPosition { get; set; }
        int SelectionStart { get; set; }
        int SelectionLength { get; set; }
        bool ShowCursor { get; set; }

        // Padding interno
        float PaddingLeft { get; set; }
        float PaddingTop { get; set; }
        float PaddingRight { get; set; }
        float PaddingBottom { get; set; }

        // Métodos
        [ComVisible(true)]
        stdole.IPictureDisp RenderTextBox();
        [ComVisible(true)]
        void SelectAll();
        [ComVisible(true)]
        void ClearSelection();
        [ComVisible(true)]
        string GetSelectedText();
        [ComVisible(true)]
        void SetSelection(int start, int length);
        [ComVisible(true)]
        void MoveCursorToEnd();
        [ComVisible(true)]
        void MoveCursorToStart();
        [ComVisible(true)]
        int GetCursorPositionFromPoint(float x, float y);
    }

    [ComVisible(true)]
    [Guid("E0E5F6F4-4567-7890-1234-456789012DEF")]
    [ClassInterface(ClassInterfaceType.None)]
    public class SkiaTextBox : ISkiaTextBox
    {
        private string _text = "";
        private string _placeholderText = "Digite aqui...";
        private int _maxLength = 0; // 0 = sem limite
        private bool _readOnly = false;
        private bool _multiline = false;
        private int _textAlign = 0; // 0=Left, 1=Center, 2=Right

        // Aparência
        private int _width = 200;
        private int _height = 32;
        private uint _backgroundColor = 0xFFFFFFFF; // Branco
        private uint _textColor = 0xFF333333; // Cinza escuro
        private uint _placeholderColor = 0xFF999999; // Cinza médio
        private uint _borderColor = 0xFFCCCCCC; // Cinza claro
        private uint _focusBorderColor = 0xFF4285F4; // Azul Google
        private uint _selectionColor = 0xFF4285F4; // Azul para seleção
        private float _borderWidth = 1.0f;
        private float _cornerRadius = 4.0f;

        // Fonte
        private string _fontFamily = "Segoe UI";
        private float _fontSize = 14.0f;
        private bool _bold = false;

        // Estado
        private bool _hasFocus = false;
        private bool _enabled = true;
        private int _cursorPosition = 0;
        private int _selectionStart = 0;
        private int _selectionLength = 0;
        private bool _showCursor = true;

        // Padding
        private float _paddingLeft = 8.0f;
        private float _paddingTop = 6.0f;
        private float _paddingRight = 8.0f;
        private float _paddingBottom = 6.0f;

        #region Propriedades COM

        public string Text
        {
            get => _text;
            set
            {
                _text = value ?? "";
                if (_cursorPosition > _text.Length)
                    _cursorPosition = _text.Length;
                ClearSelection();
            }
        }

        public string PlaceholderText
        {
            get => _placeholderText;
            set => _placeholderText = value ?? "";
        }

        public int MaxLength
        {
            get => _maxLength;
            set => _maxLength = Math.Max(0, value);
        }

        public bool ReadOnly
        {
            get => _readOnly;
            set => _readOnly = value;
        }

        public bool Multiline
        {
            get => _multiline;
            set => _multiline = value;
        }

        public int TextAlign
        {
            get => _textAlign;
            set => _textAlign = Math.Max(0, Math.Min(2, value));
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
            get => unchecked((int)_backgroundColor);
            set => _backgroundColor = unchecked((uint)value);
        }

        public int TextColor
        {
            get => unchecked((int)_textColor);
            set => _textColor = unchecked((uint)value);
        }

        public int PlaceholderColor
        {
            get => unchecked((int)_placeholderColor);
            set => _placeholderColor = unchecked((uint)value);
        }

        public int BorderColor
        {
            get => unchecked((int)_borderColor);
            set => _borderColor = unchecked((uint)value);
        }

        public int FocusBorderColor
        {
            get => unchecked((int)_focusBorderColor);
            set => _focusBorderColor = unchecked((uint)value);
        }

        public int SelectionColor
        {
            get => unchecked((int)_selectionColor);
            set => _selectionColor = unchecked((uint)value);
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

        public bool HasFocus
        {
            get => _hasFocus;
            set => _hasFocus = value;
        }

        public bool Enabled
        {
            get => _enabled;
            set => _enabled = value;
        }

        public int CursorPosition
        {
            get => _cursorPosition;
            set => _cursorPosition = Math.Max(0, Math.Min(_text.Length, value));
        }

        public int SelectionStart
        {
            get => _selectionStart;
            set => _selectionStart = Math.Max(0, Math.Min(_text.Length, value));
        }

        public int SelectionLength
        {
            get => _selectionLength;
            set => _selectionLength = Math.Max(0, Math.Min(_text.Length - _selectionStart, value));
        }

        public bool ShowCursor
        {
            get => _showCursor;
            set => _showCursor = value;
        }

        public float PaddingLeft
        {
            get => _paddingLeft;
            set => _paddingLeft = Math.Max(0, value);
        }

        public float PaddingTop
        {
            get => _paddingTop;
            set => _paddingTop = Math.Max(0, value);
        }

        public float PaddingRight
        {
            get => _paddingRight;
            set => _paddingRight = Math.Max(0, value);
        }

        public float PaddingBottom
        {
            get => _paddingBottom;
            set => _paddingBottom = Math.Max(0, value);
        }

        #endregion

        #region Métodos COM

        [ComVisible(true)]
        public stdole.IPictureDisp RenderTextBox()
        {
            SKBitmap bitmap = new SKBitmap(_width, _height);
            SKCanvas canvas = new SKCanvas(bitmap);
            canvas.Clear(SKColors.Transparent);

            // Desenhar o fundo
            DrawBackground(canvas);

            // Desenhar borda
            DrawBorder(canvas);

            // Desenhar seleção (se houver)
            if (_selectionLength > 0)
            {
                DrawSelection(canvas);
            }

            // Desenhar texto ou placeholder
            if (string.IsNullOrEmpty(_text))
            {
                DrawPlaceholder(canvas);
            }
            else
            {
                DrawText(canvas);
            }

            // Desenhar cursor (se tiver foco e estiver visível)
            if (_hasFocus && _showCursor && _enabled && !_readOnly)
            {
                DrawCursor(canvas);
            }

            return ToPicture(bitmap);
        }

        [ComVisible(true)]
        public void SelectAll()
        {
            _selectionStart = 0;
            _selectionLength = _text.Length;
            _cursorPosition = _text.Length;
        }

        [ComVisible(true)]
        public void ClearSelection()
        {
            _selectionStart = 0;
            _selectionLength = 0;
        }

        [ComVisible(true)]
        public string GetSelectedText()
        {
            if (_selectionLength > 0 && _selectionStart < _text.Length)
            {
                return _text.Substring(_selectionStart, Math.Min(_selectionLength, _text.Length - _selectionStart));
            }
            return "";
        }

        [ComVisible(true)]
        public void SetSelection(int start, int length)
        {
            _selectionStart = Math.Max(0, Math.Min(_text.Length, start));
            _selectionLength = Math.Max(0, Math.Min(_text.Length - _selectionStart, length));
            _cursorPosition = _selectionStart + _selectionLength;
        }

        [ComVisible(true)]
        public void MoveCursorToEnd()
        {
            _cursorPosition = _text.Length;
            ClearSelection();
        }

        [ComVisible(true)]
        public void MoveCursorToStart()
        {
            _cursorPosition = 0;
            ClearSelection();
        }

        [ComVisible(true)]
        public int GetCursorPositionFromPoint(float x, float y)
        {
            if (string.IsNullOrEmpty(_text))
                return 0;

            using (var paint = CreateTextPaint())
            {
                var textBounds = GetTextBounds(paint);
                var textStartX = GetTextStartX(textBounds);

                if (x <= textStartX)
                    return 0;

                // Encontrar a posição mais próxima do clique
                float currentX = textStartX;
                for (int i = 0; i < _text.Length; i++)
                {
                    var charWidth = paint.MeasureText(_text.Substring(i, 1));
                    if (x < currentX + charWidth / 2)
                        return i;
                    currentX += charWidth;
                }

                return _text.Length;
            }
        }

        #endregion

        #region Métodos de Renderização

        private void DrawBackground(SKCanvas canvas)
        {
            var rect = new SKRect(0, 0, _width, _height);

            using (var paint = new SKPaint())
            {
                paint.IsAntialias = true;
                paint.Color = new SKColor((uint)(_enabled ? _backgroundColor : ModifyColor(_backgroundColor, 0.95f)));

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

        private void DrawBorder(SKCanvas canvas)
        {
            if (_borderWidth <= 0) return;

            var rect = new SKRect(_borderWidth / 2, _borderWidth / 2,
                                _width - _borderWidth / 2, _height - _borderWidth / 2);

            using (var paint = new SKPaint())
            {
                paint.IsAntialias = true;
                paint.Style = SKPaintStyle.Stroke;
                paint.StrokeWidth = _borderWidth;

                var borderColor = _hasFocus ? _focusBorderColor : _borderColor;
                paint.Color = new SKColor((uint)(_enabled ? borderColor : ModifyColor(borderColor, 0.6f)));

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

        private void DrawSelection(SKCanvas canvas)
        {
            if (_selectionLength <= 0 || string.IsNullOrEmpty(_text))
                return;

            using (var paint = CreateTextPaint())
            {
                var textBounds = GetTextBounds(paint);
                var textStartX = GetTextStartX(textBounds);
                var textY = GetTextY(textBounds);

                // Calcular posição da seleção
                var selectionStartX = textStartX;
                if (_selectionStart > 0)
                {
                    var textBeforeSelection = _text.Substring(0, _selectionStart);
                    selectionStartX += paint.MeasureText(textBeforeSelection);
                }

                var selectedText = _text.Substring(_selectionStart, Math.Min(_selectionLength, _text.Length - _selectionStart));
                var selectionWidth = paint.MeasureText(selectedText);

                // Desenhar fundo da seleção
                var selectionRect = new SKRect(
                    selectionStartX,
                    textY - textBounds.Height,
                    selectionStartX + selectionWidth,
                    textY + textBounds.Height * 0.2f);

                using (var selectionPaint = new SKPaint())
                {
                    selectionPaint.IsAntialias = true;
                    selectionPaint.Color =  ModifyColor(_selectionColor, 1.0f).WithAlpha(80);
                    canvas.DrawRect(selectionRect, selectionPaint);
                }
            }
        }

        private void DrawText(SKCanvas canvas)
        {
            using (var paint = CreateTextPaint())
            {
                paint.Color = new SKColor((uint)(_enabled ? _textColor : ModifyColor(_textColor, 0.5f)));

                var textBounds = GetTextBounds(paint);
                var textX = GetTextStartX(textBounds);
                var textY = GetTextY(textBounds);

                canvas.DrawText(_text, textX, textY, paint);
            }
        }

        private void DrawPlaceholder(SKCanvas canvas)
        {
            if (string.IsNullOrEmpty(_placeholderText))
                return;

            using (var paint = CreateTextPaint())
            {
                paint.Color = new SKColor((uint)(_enabled ? _placeholderColor : ModifyColor(_placeholderColor, 0.5f)));

                var textBounds = new SKRect();
                paint.MeasureText(_placeholderText, ref textBounds);

                var textX = _paddingLeft;
                var textY = (_height - textBounds.Height) / 2 - textBounds.Top;

                canvas.DrawText(_placeholderText, textX, textY, paint);
            }
        }

        private void DrawCursor(SKCanvas canvas)
        {
            if (string.IsNullOrEmpty(_text) && _cursorPosition == 0)
            {
                // Cursor no início quando não há texto
                var cursorX = _paddingLeft;
                var cursorY1 = _paddingTop;
                var cursorY2 = _height - _paddingBottom;

                using (var paint = new SKPaint())
                {
                    paint.Color = new SKColor(_textColor);
                    paint.StrokeWidth = 1.0f;
                    canvas.DrawLine(cursorX, cursorY1, cursorX, cursorY2, paint);
                }
            }
            else if (!string.IsNullOrEmpty(_text))
            {
                using (var paint = CreateTextPaint())
                {
                    var textBounds = GetTextBounds(paint);
                    var textStartX = GetTextStartX(textBounds);
                    var textY = GetTextY(textBounds);

                    var cursorX = textStartX;
                    if (_cursorPosition > 0)
                    {
                        var textBeforeCursor = _text.Substring(0, _cursorPosition);
                        cursorX += paint.MeasureText(textBeforeCursor);
                    }

                    using (var cursorPaint = new SKPaint())
                    {
                        cursorPaint.Color = new SKColor(_textColor);
                        cursorPaint.StrokeWidth = 1.0f;
                        canvas.DrawLine(cursorX, textY - textBounds.Height,
                                      cursorX, textY + textBounds.Height * 0.2f, cursorPaint);
                    }
                }
            }
        }

        #endregion

        #region Métodos Auxiliares

        private SKPaint CreateTextPaint()
        {
            var paint = new SKPaint();
            paint.IsAntialias = true;
            paint.TextSize = _fontSize;
            paint.Typeface = SKTypeface.FromFamilyName(_fontFamily,
                _bold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal,
                SKFontStyleWidth.Normal,
                SKFontStyleSlant.Upright);
            return paint;
        }

        private SKRect GetTextBounds(SKPaint paint)
        {
            var bounds = new SKRect();
            var displayText = string.IsNullOrEmpty(_text) ? _placeholderText : _text;
            if (!string.IsNullOrEmpty(displayText))
            {
                paint.MeasureText(displayText, ref bounds);
            }
            return bounds;
        }

        private float GetTextStartX(SKRect textBounds)
        {
            switch (_textAlign)
            {
                case 1: // Center
                    return (_width - textBounds.Width) / 2;
                case 2: // Right
                    return _width - _paddingRight - textBounds.Width;
                default: // Left
                    return _paddingLeft;
            }
        }

        private float GetTextY(SKRect textBounds)
        {
            return (_height - textBounds.Height) / 2 - textBounds.Top;
        }

        private SKColor ModifyColor(uint color, float factor)
        {
            var skColor = new SKColor(color);
            var r = (byte)Math.Min(255, (int)(skColor.Red * factor));
            var g = (byte)Math.Min(255, (int)(skColor.Green * factor));
            var b = (byte)Math.Min(255, (int)(skColor.Blue * factor));
            return new SKColor(r, g, b, skColor.Alpha);
        }

        private IPictureDisp ToPicture(SKBitmap bitmap)
        {
            SKImage image = SKImage.FromPixels(bitmap.PeekPixels());
            SKData encoded = image.Encode();
            Stream stream = encoded.AsStream();
            var img = Image.FromStream(stream);
            return ImageToIPictureConverter.ConvertImageToIPictureDisp(img);
        }

        #endregion
    }
}