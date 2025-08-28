using stdole;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SkiaSharpCOM
{
    public class ImageToIPictureConverter : AxHost
    {
        // Construtor necessário - pode usar qualquer CLSID válido
        public ImageToIPictureConverter() : base("59CCB4A0-729D-11CE-8C43-00AA004CD6D8")
        {

        }

        // Método público que expõe o método protegido
        public static IPictureDisp ConvertImageToIPictureDisp(Image image)
        {
            if (image == null) return null;

            return (IPictureDisp)GetIPictureDispFromPicture(image);
        }

        // Método para converter de arquivo diretamente
        public static IPictureDisp LoadImageFileToIPictureDisp(string filePath)
        {
            if (!System.IO.File.Exists(filePath))
                throw new System.IO.FileNotFoundException("Arquivo não encontrado: " + filePath);

            using (Image image = Image.FromFile(filePath))
            {
                return ConvertImageToIPictureDisp(image);
            }
        }

        // Método para converter de stream
        public static IPictureDisp ConvertStreamToIPictureDisp(System.IO.Stream stream)
        {
            if (stream == null) return null;

            using (Image image = Image.FromStream(stream))
            {
                return ConvertImageToIPictureDisp(image);
            }
        }
    }
}
