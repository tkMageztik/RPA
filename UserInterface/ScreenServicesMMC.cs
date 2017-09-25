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
    class ScreenServicesMMC : IUserInterface
    {
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
            //Process[] p = Process.GetProcessesByName("mmc");

            //System.Windows.Automation.AutomationElement t = WinFormAdapter.GetAEFromHandle(p[0].MainWindowHandle);
            //System.Windows.Automation.AutomationElement _0 = AutomationElement.FromHandle(p[0].MainWindowHandle);

            //System.Windows.Automation.AutomationElement t1 = WinFormAdapter.GetAEOnChildByName(t, "Sesión A - [24 x 80]");
            //System.Windows.Automation.AutomationElement t1 = t.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Inicio de sesión de IBM i"));

            //TODO: Obtener por tipo de elemento (Tipo.Edit)
            //Al tener el mismo nombre tanto el label como el textbox de un campo, toma los dos, al parecer el segundo siempre es el textbox....

            Process px = Process.Start(@"C:\Windows\System32\calc.exe");
            

            AutomationElement aeDesktop = AutomationElement.RootElement;
            Process[] p = Process.GetProcessesByName("calc");

            //System.Windows.Automation.AutomationElement t = WinFormAdapter.GetAEFromHandle(p[0].MainWindowHandle);
            System.Windows.Automation.AutomationElement aeCalculator = AutomationElement.FromHandle(p[0].MainWindowHandle);


            //AutomationElement aeCalculator;



            int numwaits = 0;
            do
            {
                Debug.WriteLine("Looking for Calculator . . . ");
                //aeCalculator = aeDesktop.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Calculadora"));
                numwaits += 1;
                //Thread.Sleep(100)

            } while (numwaits == null && numwaits < 50);




            String btn5hexID = "00000087";
            String btn5decimalID = Convert.ToInt32("00000087", 16).ToString();

            String btnAddhexID = "0000005D";
            String btnAdddecimalID = Convert.ToInt32("0000005D", 16).ToString();

            String btnEqualshexID = "00000079";
            String btnEqualsdecimalID = Convert.ToInt32("00000079", 16).ToString();

            AutomationElement ae5Btn = aeCalculator.FindFirst(TreeScope.Descendants,new PropertyCondition(AutomationElement.AutomationIdProperty, btn5decimalID));

            AutomationElement aeAddBtn = aeCalculator.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, btnAdddecimalID));

            AutomationElement aeEqualsBtn = aeCalculator.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, btnEqualsdecimalID));



            InvokePattern ipClick5Btn = (InvokePattern)ae5Btn.GetCurrentPattern(InvokePattern.Pattern);
            InvokePattern ipClickAddBtn = (InvokePattern)aeAddBtn.GetCurrentPattern(InvokePattern.Pattern);
            InvokePattern ipClickEqualsBtn = (InvokePattern)aeEqualsBtn.GetCurrentPattern(InvokePattern.Pattern);
            aeCalculator.SetFocus();
            ipClick5Btn.Invoke();
            ipClickAddBtn.Invoke();
            ipClick5Btn.Invoke();
            ipClickEqualsBtn.Invoke();




            /*

            System.Windows.Automation.AutomationElementCollection x = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Servicios (locales)"));

            System.Windows.Automation.AutomationElementCollection xx = x[0].FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Web Teller -  Ruteos "));
            System.Windows.Automation.AutomationElementCollection x2 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Web Teller -  Ruteos "));
            System.Windows.Automation.AutomationElementCollection x3 = _0.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Web Teller -  Ruteos "));
            System.Windows.Automation.AutomationElementCollection x4 = _0.FindAll(TreeScope.Element, new PropertyCondition(AutomationElement.NameProperty, "Web Teller -  Ruteos "));

            System.Windows.Automation.AutomationElement y = WinFormAdapter.GetAEOnDescByName(x[0], ControlType.Edit, "Web Teller -  Ruteos ");
            System.Windows.Automation.AutomationElement y1 = WinFormAdapter.GetAEOnDescByName(xx[0], ControlType.Edit, "Web Teller -  Ruteos ");


            InvokePattern eee = (InvokePattern)y.GetCurrentPattern(InvokePattern.Pattern);
            InvokePattern eee2 = (InvokePattern)y1.GetCurrentPattern(InvokePattern.Pattern);


            eee.Invoke();
            eee2.Invoke();
            
            System.Windows.Automation.AutomationElement y2 = WinFormAdapter.GetAEOnDescByName(xx[1], ControlType.Edit, "Web Teller -  Ruteos ");
            System.Windows.Automation.AutomationElement y3 = WinFormAdapter.GetAEOnDescByName(xx[2], ControlType.Edit, "Web Teller -  Ruteos ");
            System.Windows.Automation.AutomationElement y4 = WinFormAdapter.GetAEOnDescByName(x3[0], ControlType.Edit, "Web Teller -  Ruteos ");
            System.Windows.Automation.AutomationElement y5 = WinFormAdapter.GetAEOnDescByName(x3[1], ControlType.Edit, "Web Teller -  Ruteos ");
            System.Windows.Automation.AutomationElement y6 = WinFormAdapter.GetAEOnDescByName(x3[2], ControlType.Edit, "Web Teller -  Ruteos ");




            WinFormAdapter.ClickElement(y);
            WinFormAdapter.ClickElement(y1);

            WinFormAdapter.ClickElement(xx[2]);
            WinFormAdapter.ClickElement(xx[1]);
            WinFormAdapter.ClickElement(xx[0]);

            */
            //x[0].FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, ""));


            /*System.Windows.Automation.AutomationElementCollection x1 = _0.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Servicios (locales)"));

            System.Windows.Automation.AutomationElementCollection x2 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Área de trabajo"));
            System.Windows.Automation.AutomationElementCollection x3 = _0.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Área de trabajo"));
            */

            /*System.Windows.Automation.AutomationElementCollection _0_Descendants_1 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Servicios"));

            System.Windows.Automation.AutomationElementCollection _0_Childs_1 = _0.FindAll(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Servicios"));


            WinFormAdapter.SetText(_0_Descendants_1[1], "BFPJUARUI");

            System.Windows.Automation.AutomationElementCollection _0_Descendants_2 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Contraseña:"));
            WinFormAdapter.SetText(_0_Descendants_2[1], "BFPJUARUI2");

            WinFormAdapter.ClickElement(WinFormAdapter.GetAEOnDescByName(_0, "Aceptar"));*/
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
