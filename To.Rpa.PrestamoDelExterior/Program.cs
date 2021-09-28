using NS.RPA.RACPAutorizacion;
using System.Configuration;

namespace To.Rpa.PrestamoDelExterior
{
    public class Program
    {
        static void Main(string[] args)
        {
            if (ConfigurationManager.AppSettings["AuthorizationPath"].Equals("SI"))
            {
                AppRACPAutorizacion appAuto = new AppRACPAutorizacion();
                appAuto.DoActivities();
            }
            else
            {
                AppPrestamoDelExterior app = new AppPrestamoDelExterior();
                app.DoActivities();
            }
        }
    }
}
