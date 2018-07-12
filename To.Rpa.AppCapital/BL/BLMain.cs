using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using To.Rpa.AppCapital.Interfaces;
using To.Rpa.AppCapital.UserInterfaces;
using To.AtNinjas.Util;

namespace To.Rpa.AppCapital.BL
{
    public class BLMain
    {
        public void DoActivities()
        {
            Login();
        }

        public void Login()
        {
            try
            {

                if (Convert.ToBoolean(ConfigurationManager.AppSettings["DualValidation"].ToString().Trim()))
                {
                    new ScreenPreviousLogin().DoActivities();
                }

                new ScreenLogin().DoActivities();
            }
            catch (Exception exc)
            {
                throw exc;
            }
        }
    }
}
