using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace To.Rpa.Util
{
    public class Methods
    {
        public static void LogProceso(string sMensaje)
        {
            string dirLog = ConfigurationManager.AppSettings["DirLog"].ToString().Trim();
            string ArcExt = ConfigurationManager.AppSettings["ArcExt"].ToString().Trim();

            DateTime fechaHora = DateTime.Now;
            string fileName = String.Format("{0:ddMMyyyy}", fechaHora);

            string path = dirLog + fileName + ArcExt;
            try
            {
                if (File.Exists(path))
                {
                    File.AppendAllText(path, DateTime.Now.ToString() + " | " + sMensaje.Trim() + " \r\n");
                }
                else
                {
                    using (var file = File.Create(path))
                    {
                        file.Close();
                    }
                    File.AppendAllText(path, DateTime.Now.ToString() + " | " + sMensaje.Trim() + " \r\n");

                }
            }
            catch { }
        }
        //TODO: en parametro estaba this, es método de extensión, investigar
        public static Bitmap cropAtRect(Bitmap b, Rectangle r)
        {
            Bitmap nb = new Bitmap(r.Width, r.Height);
            Graphics g = Graphics.FromImage(nb);
            g.DrawImage(b, -r.X, -r.Y);
            return nb;
        }

        public static void Sleep()
        {
            Thread.Sleep(Convert.ToInt32(ConfigurationManager.AppSettings["Sleep"]));
        }
    }
}
