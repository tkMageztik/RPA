using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using To.AtNinjas.UserControls;
using To.AtNinjas.Util;

namespace To.Rpa.AppCTSSA
{
    public partial class FrmReader : Form
    {
        bool bHaveMouse;
        Point ptOriginal = new Point();
        Point ptLast = new Point();
        Rectangle rectCropArea;

        public FrmReader()
        {
            InitializeComponent();
            bHaveMouse = false;
        }

        private void FrmReader_Load(object sender, EventArgs e)
        {
            string imagePath = @"d:\Users\juarui\Source\Repos\RPA\To.Rpa.AppCTS\CTS\test\4023\base_4023_0.png";
            //imagePath = @"d: \Users\juarui\Desktop\Mobilex proposal\crop example\demos_utiles_karen_standalone\selector area\C#\CSWinformCropImage\1.jpeg";

            //SrcPicBox.Image = Image.FromFile(imagePath);
            imagePanel1.Image = new Bitmap(imagePath);
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
                imagePanel1.Refresh();
            }
        }

        private void SrcPicBox_Paint(object sender, PaintEventArgs e)
        {
            Pen drawLine = new Pen(Color.Black);
            drawLine.DashStyle = DashStyle.Dash;
            //e.Graphics.DrawRectangle(drawLine, rectCropArea);


            Graphics.FromImage(((ImagePanel)sender).Image).DrawRectangle(drawLine, rectCropArea);


        }

        private void Crop()
        {
            //TargetPicBox.Refresh();
            //Prepare a new Bitmap on which the cropped image will be drawn
            Bitmap sourceBitmap = new Bitmap(imagePanel1.Image);
            //Graphics g = TargetPicBox.CreateGraphics();

            Bitmap newImage = Methods.cropAtRect(sourceBitmap, rectCropArea);

            //Checks if the co-rdinates check-box is checked. If yes, then Selection is based on co-rdinates mentioned in the textbox
            //if (chkCropCordinates.Checked)
            //{
            //    //logic to retreive co-rdinates from comma-separated string values
            //    //lbCordinates.Text = "";
            //    //string[] cordinates = tbCordinates.Text.ToString().Split(',');
            //    int cord0, cord1, cord2, cord3;

            //    try
            //    {
            //        cord0 = Convert.ToInt32(cordinates[0]);
            //        cord1 = Convert.ToInt32(cordinates[1]);
            //        cord2 = Convert.ToInt32(cordinates[2]);
            //        cord3 = Convert.ToInt32(cordinates[3]);
            //    }
            //    catch (Exception ex)
            //    {
            //        //lbCordinates.Text = ex.Message;
            //        return;
            //    }

            //    //Various combinations of selection rectangle being dragged in different directions

            //    if ((cord0 < cord2 && cord1 < cord3))
            //    {
            //        rectCropArea = new Rectangle(cord0, cord1, cord2 - cord0, cord3 - cord1);
            //    }
            //    else if (cord2 < cord0 && cord3 > cord1)
            //    {
            //        rectCropArea = new Rectangle(cord2, cord1, cord0 - cord2, cord3 - cord1);
            //    }
            //    else if (cord2 > cord0 && cord3 < cord1)
            //    {
            //        rectCropArea = new Rectangle(cord0, cord3, cord2 - cord0, cord1 - cord3);
            //    }
            //    else
            //    {
            //        rectCropArea = new Rectangle(cord2, cord3, cord0 - cord2, cord1 - cord3);
            //    }
            //}



            //Draw the image on the Graphics object with the new dimesions
            //g.DrawImage(sourceBitmap, new Rectangle(0, 0, TargetPicBox.Width, TargetPicBox.Height),
            //    rectCropArea, GraphicsUnit.Pixel);

            ////Good practice to dispose the System.Drawing objects when not in use.
            //sourceBitmap.Dispose();

            newImage.Save(@"C:\test.png");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Crop();
        }
    }
}
