using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tcs.Rpa.AppCapital.Mainframe.Interfaces;
using Tcs.Rpa.AppCapital.Mainframe.UserInterfaces;

namespace Tcs.Rpa.AppCapital.Mainframe.BL
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
                new ScreenPreviousLogin().DoActivities();
                //new ScreenLogin().DoActivities();
            }
            catch (Exception exc) { }
        }
    }
}
