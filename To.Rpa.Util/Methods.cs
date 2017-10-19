using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Tcs.Rpa.Util
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
    }
}
