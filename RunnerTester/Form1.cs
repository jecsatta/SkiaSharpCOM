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
            SkiaSharpCOM.SkiaRenderer renderer = new SkiaSharpCOM.SkiaRenderer();
            renderer.CreateBitmap(200, 200);
            renderer.DrawCircle(50, 50, 25, 0, 2, true);

            pictureBox1.Image = Microsoft.VisualBasic.Compatibility.VB6.Support.IPictureDispToImage(renderer.ToPicture());


        }
    }
}
