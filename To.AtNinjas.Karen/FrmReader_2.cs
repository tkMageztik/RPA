using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;
using To.AtNinjas.Util;

namespace To.AtNinjas.Karen
{
    public partial class FrmReader_2 : Form
    {
        bool bHaveMouse;
        Point ptOriginal = new Point();
        Point ptLast = new Point();
        Rectangle rectCropArea;
        private string AppTempDirectory { get; set; }

        public FrmReader_2()
        {
            AppTempDirectory = ConfigurationManager.AppSettings["AppTempDirectory"];
            InitializeComponent();
            bHaveMouse = false;
        }

        private void FrmReader_Load(object sender, EventArgs e)
        {
            string imagePath = @"d:\Users\juarui\Source\Repos\RPA\To.Rpa.AppCTS\CTS\test\4023\base_4023_0.png";
            //imagePath = @"d: \Users\juarui\Desktop\Mobilex proposal\crop example\demos_utiles_karen_standalone\selector area\C#\CSWinformCropImage\1.jpeg";

            SrcPicBox.Image = Image.FromFile(imagePath);
            //imagePanel1.Image = new Bitmap(imagePath);
        }

        private void SrcPicBox_MouseDown(object sender, MouseEventArgs e)
        {
            // Make a note that we "have the mouse".
            bHaveMouse = true;

            // Store the "starting point" for this rubber-band rectangle.
            ptOriginal.X = e.X;
            ptOriginal.Y = e.Y;

            // Special value lets us know that no previous
            // rectangle needs to be erased.

            // Display coordinates
            //lbCordinates.Text = "Coordinates  :  " + e.X.ToString() + ", " + e.Y.ToString();

            ptLast.X = -1;
            ptLast.Y = -1;

            rectCropArea = new Rectangle(new Point(e.X, e.Y), new Size());
        }

        private void SrcPicBox_MouseUp(object sender, MouseEventArgs e)
        {
            // Set internal flag to know we no longer "have the mouse".
            bHaveMouse = false;

            // If we have drawn previously, draw again in that spot
            // to remove the lines.
            if (ptLast.X != -1)
            {
                Point ptCurrent = new Point(e.X, e.Y);

                // Display coordinates
                //lbCordinates.Text = "Coordinates  :  " + ptOriginal.X.ToString() + ", " +
                //    ptOriginal.Y.ToString() + " And " + e.X.ToString() + ", " + e.Y.ToString();

            }

            // Set flags to know that there is no "previous" line to reverse.
            ptLast.X = -1;
            ptLast.Y = -1;
            ptOriginal.X = -1;
            ptOriginal.Y = -1;

        }

        private void SrcPicBox_MouseMove(object sender, MouseEventArgs e)
        {
            Point ptCurrent = new Point(e.X, e.Y);

            // If we "have the mouse", then we draw our lines.
            if (bHaveMouse)
            {
                // If we have drawn previously, draw again in
                // that spot to remove the lines.
                //if (ptLast.X != -1)
                //{
                //    // Display Coordinates
                //    lbCordinates.Text = "Coordinates  :  " + ptOriginal.X.ToString() + ", " +
                //        ptOriginal.Y.ToString() + " And " + e.X.ToString() + ", " + e.Y.ToString();
                //}

                // Update last point.
                ptLast = ptCurrent;

                // Draw new lines.

                // e.X - rectCropArea.X;
                // normal
                if (e.X > ptOriginal.X && e.Y > ptOriginal.Y)
                {
                    rectCropArea.Width = e.X - ptOriginal.X;

                    // e.Y - rectCropArea.Height;
                    rectCropArea.Height = e.Y - ptOriginal.Y;
                }
                else if (e.X < ptOriginal.X && e.Y > ptOriginal.Y)
                {
                    rectCropArea.Width = ptOriginal.X - e.X;
                    rectCropArea.Height = e.Y - ptOriginal.Y;
                    rectCropArea.X = e.X;
                    rectCropArea.Y = ptOriginal.Y;
                }
                else if (e.X > ptOriginal.X && e.Y < ptOriginal.Y)
                {
                    rectCropArea.Width = e.X - ptOriginal.X;
                    rectCropArea.Height = ptOriginal.Y - e.Y;

                    rectCropArea.X = ptOriginal.X;
                    rectCropArea.Y = e.Y;
                }
                else
                {
                    rectCropArea.Width = ptOriginal.X - e.X;

                    // e.Y - rectCropArea.Height;
                    rectCropArea.Height = ptOriginal.Y - e.Y;
                    rectCropArea.X = e.X;
                    rectCropArea.Y = e.Y;
                }
                SrcPicBox.Refresh();
            }
        }

        private void SrcPicBox_Paint(object sender, PaintEventArgs e)
        {
            Pen drawLine = new Pen(Color.Black);
            drawLine.DashStyle = DashStyle.Dash;
            e.Graphics.DrawRectangle(drawLine, rectCropArea);


            //Graphics.FromImage(((ImagePanel)sender).Image).DrawRectangle(drawLine, rectCropArea);


        }

        private void Crop()
        {
            if (rectCropArea == Rectangle.Empty) return;

            Bitmap sourceBitmap = new Bitmap(SrcPicBox.Image);

            double xfactor = Convert.ToDouble(SrcPicBox.Image.Width) / SrcPicBox.Size.Width;
            double yfactor = Convert.ToDouble(SrcPicBox.Image.Height) / SrcPicBox.Size.Height;

            rectCropArea.X = Convert.ToInt32(rectCropArea.X * xfactor);
            rectCropArea.Y = Convert.ToInt32(rectCropArea.Y * yfactor);
            rectCropArea.Height = Convert.ToInt32(rectCropArea.Height * yfactor);
            rectCropArea.Width = Convert.ToInt32(rectCropArea.Width * xfactor);

            Bitmap newImage = Methods.cropAtRect(sourceBitmap, rectCropArea);

            newImage.Save(@"C:\temp.png");
            rectCropArea = Rectangle.Empty;
            sourceBitmap.Dispose();
            newImage.Dispose();
            SrcPicBox.Refresh();

            List<List<string>> lst = new KarenCore().GetData(AppTempDirectory);

            for (int i = 0; i <= lst.Count; i++)
            {
                List<string> lst_2 = lst[i];

                string column = "";
                switch (i)
                {
                    case 0: column = "A"; break;
                    case 1: column = "B"; break;
                    case 2: column = "C"; break;
                    case 3: column = "D"; break;
                    case 4: column = "E"; break;
                    case 5: column = "F"; break;
                    case 6: column = "G"; break;
                    case 7: column = "H"; break;
                    case 8: column = "I"; break;
                    case 9: column = "J"; break;
                    default: column = "K"; break;
                }

                for (int x = 0; x <= lst_2.Count; x++)
                {
                    try
                    {

                    }
                    catch (Exception exc) { }
                }
            }

            //foreach (List<string> lst_2 in lst)
            //{
            //    foreach (string str in lst_2)
            //    {
            //        try
            //        {
            //            Methods.UpdateCell("", "", 0, "");
            //        }
            //        catch (Exception exc) { }
            //    }
            //}
        }





        private void button1_Click(object sender, EventArgs e)
        {
            Crop();
        }

        private void SrcPicBox_Click(object sender, EventArgs e)
        {
            SrcPicBox.Refresh();
        }
    }
}
