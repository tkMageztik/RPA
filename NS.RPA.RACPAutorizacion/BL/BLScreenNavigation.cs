using EHLLAPI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using To.AtNinjas.Util;

namespace NS.RPA.RACPAutorizacion.BL
{
    class BLScreenNavigation
    {
        CustomEHLLAPI ehllapi;
        public BLScreenNavigation()
        {
            ehllapi = new CustomEHLLAPI();
            // ehllapi.Connect(SessionId);
        }

        public string ReadScreen(string position, int lenght)
        {
            return "";
            //return eh.ReadScreen(position + ",1", lenght);
        }



        public bool ShowScreenCPApproval(string loanNumber)
        {
            return true;
        }
        public bool ShowScreenRATransactionsApproval(string loanNumber)
        {

            //#if DEBUG
            //            EhllapiWrapper.Wait();
            //            ehllapi.SetCursorPos("20,7");
            //            EhllapiWrapper.Wait();
            //            ehllapi.SendStr("GO MGCIA2");
            //            EhllapiWrapper.Wait();
            //            ehllapi.SendStr("@E");
            //            EhllapiWrapper.Wait();
            //#endif


            string opt = ehllapi.ReadScreen("5,3", 25).Trim();
            if (opt.Equals("1. Posición de cliente"))
            {
                Methods.LogProceso("El menú inicial es el definido: MGCIA2");
            }
            else
            {
                Methods.LogProceso("El menú inicial no es el definido, se aborta el RPA");
                return false;
            }

            EhllapiWrapper.Wait();
            Methods.GetMenuProgram("15,1");
            EhllapiWrapper.Wait();

            if (!ehllapi.ReadScreen("2,20", 41).Contains("Aprobacion de Contratos y Transacciones"))
            {
                Methods.LogProceso("No se abrió correctamente la opción 'Aprobacion de Contratos y Transacciones'");
                return false;
            }
            return true;
        }

        public void BackToMainMenu()
        {

            ehllapi.SendStr("@v");

            ehllapi.SendStr("@7");
            EhllapiWrapper.Wait();
            ehllapi.SendStr("@7");
            EhllapiWrapper.Wait();
            ehllapi.SendStr("@7");
            EhllapiWrapper.Wait();
        }
    }
}
