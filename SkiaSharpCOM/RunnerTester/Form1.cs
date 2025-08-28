using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RunnerTester
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

           /*  SkiaSharpCOM.SkiaRenderer renderer = new SkiaSharpCOM.SkiaRenderer();
             renderer.CreateBitmap(200, 200);
             renderer.DrawCircle(50, 50, 25, 0, 2, true);

             pictureBox1.Image = Microsoft.VisualBasic.Compatibility.VB6.Support.IPictureDispToImage(renderer.ToPicture());
            
            */
            SkiaSharpCOM.SkiaButton btnPrimary = new SkiaSharpCOM.SkiaButton();

            btnPrimary.Text = "Primary";
            btnPrimary.Width = 180;
            btnPrimary.Height = 50;

            btnPrimary.BorderWidth = 0;
            btnPrimary.CornerRadius = 6;
            btnPrimary.FontFamily = "Segoe UI";
            btnPrimary.FontSize = 14;
            btnPrimary.Bold = false;
            pictureBox1.Image = Microsoft.VisualBasic.Compatibility.VB6.Support.IPictureDispToImage(btnPrimary.RenderButton());
            
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
