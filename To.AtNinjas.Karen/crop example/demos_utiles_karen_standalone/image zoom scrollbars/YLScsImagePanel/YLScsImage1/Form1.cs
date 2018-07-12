using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace YLScsImage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void imagePanel1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Bitmap b = new Bitmap(openFileDialog1.FileName);
               //imagePanel1.CanvasSize = b.Size;
                imagePanel1.Image = b;
            }
        }

        private void trackBar1_Scroll(object sender, EventArgs e)
        {
            imagePanel1.Zoom = trackBar1.Value * 0.02f;
        }
    }
}