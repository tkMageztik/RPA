using MainframeAutomationSample.HelperClasses;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Automation;
using TCSRPA;

namespace MainframeAutomationSample.UserInterface
{
    class ScreenPreviousLogin : IUserInterface
    {
        private Stopwatch objWatch = new Stopwatch();

        public void CheckCompleteLoad()
        {

        }

        public void DoActivities()
        {
            EntityGlobal.lastScreenName = "PreviousLogin";

            CheckCompleteLoad();
            // ExtractElements();
            // DoUIValidation();
            DoOperations();
        }

        public void DoOperations()
        {
            Process[] p = Process.GetProcessesByName("pcsws");

            //System.Windows.Automation.AutomationElement t = WinFormAdapter.GetAEFromHandle(p[0].MainWindowHandle);
            System.Windows.Automation.AutomationElement _0 = AutomationElement.FromHandle(p[0].MainWindowHandle);

            //System.Windows.Automation.AutomationElement t1 = WinFormAdapter.GetAEOnChildByName(t, "Sesión A - [24 x 80]");
            //System.Windows.Automation.AutomationElement t1 = t.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Inicio de sesión de IBM i"));

            //TODO: Obtener por tipo de elemento (Tipo.Edit)
            //Al tener el mismo nombre tanto el label como el textbox de un campo, toma los dos, al parecer el segundo siempre es el textbox....
            System.Windows.Automation.AutomationElementCollection _0_Descendants_1 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "ID de usuario:"));
            WinFormAdapter.SetText(_0_Descendants_1[1], "BFPJUARUI");

            System.Windows.Automation.AutomationElementCollection _0_Descendants_2 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Contraseña:"));
            WinFormAdapter.SetText(_0_Descendants_2[1], "BFPJUARUI2");

            WinFormAdapter.ClickElement(WinFormAdapter.GetAEOnDescByName(_0, "Aceptar"));
        }

        public void DoUIValidation()
        {
            throw new NotImplementedException();
        }

        public void ExtractElements()
        {
            throw new NotImplementedException();
        }
    }
}
