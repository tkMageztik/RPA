using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tcs.Rpa.AppCapital.Mainframe.BL;
using Tcs.Rpa.Util;

namespace Tcs.Rpa.AppCapital.Mainframe
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());

            //TODO: Para configurar máximos intentos
            int i = 0;
            int attempts = 3;

            do
            {
                try
                {
                    Methods.LogProceso(" Se inicia proceso de AppCapital" + MethodBase.GetCurrentMethod().Name);

                    Process px = Process.Start(@"C:\Program Files\IBM\Client Access\Emulator\Private\LAMBDA.WS");

                    Process[] p = Process.GetProcessesByName("pcsws");

                    new BLMain().DoActivities();

                }

                catch (Exception exc)
                {
                    i++;
                    Methods.LogProceso(" ERROR: " + exc.StackTrace + " " + MethodBase.GetCurrentMethod().Name);
                }
                finally
                {
                    Methods.LogProceso(" FINALLY: Intento nro." + attempts);
                }

            } while (i < attempts);
        }

    }
}
