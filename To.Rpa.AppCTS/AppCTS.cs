using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using To.Rpa.Util;
using To.Rpa.AppCTS.BL;

namespace To.Rpa.AppCTS
{
    public class AppCTS
    {
        public AppCTS()
        {
            //TODO: Para configurar máximos intentos
            UInt16 attempts = Convert.ToUInt16(ConfigurationManager.AppSettings["Attemps"]);
            int i = 0;
            do
            {
                try
                {


                    Assembly assembly = Assembly.GetExecutingAssembly();
                    FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);

                    //TODO: La configuración puede estar a nivel de la torre de control
                    Methods.LogProceso("Se inicia proceso de AppCTS: " + fvi.ProductName + " -> " + MethodBase.GetCurrentMethod().Name);

                    new BLMain().DoActivities();

                }
                catch (Exception exc)
                {
                    i++;
                    Methods.LogProceso(" ERROR: " + exc.StackTrace + " " + MethodBase.GetCurrentMethod().Name);
                    //CloseEmulator();
                }
                finally
                {
                    Methods.LogProceso(" FINALLY: Intento nro: " + i);
                }
            }
            while (i < attempts);
        }
    }
}
