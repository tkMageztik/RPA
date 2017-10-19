using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;

namespace Tcs.Rpa.AppCapital.Mainframe.Interfaces
{
    public class ScreenPreviousLogin : IUserInterface
    {
        public void CheckCompleteLoad()
        {
        }

        public void DoActivities()
        {
            CheckCompleteLoad();
            // ExtractElements();
            // DoUIValidation();
            DoOperations();
        }

        public void DoOperations()
        {
            Process[] p = Process.GetProcessesByName("pcsws");
            //p[0].WaitForInputIdle();

            Thread.Sleep(2000);

            AutomationElement _0 = AutomationElement.FromHandle(p[0].MainWindowHandle);

            AutomationElementCollection _0_Descendants_1 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "ID de usuario:"));
            //WinFormAdapter.SetText(_0_Descendants_1[1], "BFPJUARUI");

            ValuePattern etb = _0_Descendants_1[1].GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
            etb.SetValue("BFPJUARUI");           
            
            AutomationElementCollection _0_Descendants_2 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Contraseña:"));
            //WinFormAdapter.SetText(_0_Descendants_2[1], "BFPJUARUI2");

            ValuePattern etb2 = _0_Descendants_2[1].GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
            etb2.SetValue("BFPJUARUI3");

            //WinFormAdapter.ClickElement(WinFormAdapter.GetAEOnDescByName(_0, "Aceptar"));
            AutomationElementCollection _0_Descendants_3 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Aceptar"));

            var invokePattern = _0_Descendants_3[0].GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            invokePattern.Invoke();

            //intento 1
            //var tt = _0_Descendants_3[0].GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
            //invokePattern.Invoke();

            //intento 2
            //AutomationElementCollection xx = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Sesión A - [24 x 80]"));

            //AutomationElement w = xx[0].FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.AutomationIdProperty, "2"));
            //AutomationElement w1 = xx[0].FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, "2"));


            //AutomationProperty[] t = xx[0].GetSupportedProperties();
            
            //TODO: Depende de que la sesión "A" esté libre


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
