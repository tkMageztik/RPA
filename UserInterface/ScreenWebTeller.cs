using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using TCSRPA;

namespace MainframeAutomationSample.UserInterface
{
    class ScreenWebTeller
    {
        public void DoActivities()
        {
            //EntityGlobal.lastScreenName = "ScreenWebMailTCS";
            //CheckCompleteLoad();
            // ExtractElements();
            // DoUIValidation();
            DoOperations();
        }

        private void DoOperations()
        {
            InternetExplorerOptions ieOptions = new InternetExplorerOptions();
            ieOptions.IntroduceInstabilityByIgnoringProtectedModeSettings = true;

            //IWebDriver driverIE = new InternetExplorerDriver(@"D:\COMPARTIDO_PUBLICO\RPA\TCS RPA V1.2 - Training\Sample\Automation - Proceso Cese - JJRDC\Mainframe Automation\Mainframe Automation Solution\packages\Selenium.WebDriver.IEDriver.3.5.1\driver", ieOptions);

            IWebDriver driverIE = new InternetExplorerDriver(System.Configuration.ConfigurationManager.AppSettings["IEDriverPath"], ieOptions);
            //driverIE.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(60);
            driverIE.Navigate().GoToUrl("http://svaplbfdes01:7777/Entorno/Inicio.aspx");

            driverIE.FindElement(By.Name("txtUsuario")).SendKeys("BFPJUARUI");
            driverIE.FindElement(By.Name("txtUsuarioFinesse")).SendKeys("003");
            driverIE.FindElement(By.Name("txtClave")).SendKeys("PRUEBAS4");

            driverIE.FindElement(By.Id("btnOK")).SendKeys(Keys.Enter);

            Thread.Sleep(3000);

            driverIE.FindElement(By.Name("txtCRMDI")).SendKeys("46786186");
            driverIE.FindElement(By.Id("btnCRMIniciar")).SendKeys(Keys.Enter);

            Thread.Sleep(3000);
            Actions acts = new Actions(driverIE);
            //driverIE.FindElement(By.Id("btnCRMIniciar")).SendKeys(Keys.Enter);

            // acts.KeyDown(Keys.Control).KeyDown(Keys.LeftAlt);
            acts.SendKeys(Keys.Enter);
            acts.Build().Perform();

            driverIE.FindElement(By.Id("buscar")).SendKeys("7100");
            acts.SendKeys(Keys.Enter);
            acts.Build().Perform();

            driverIE.FindElement(By.Name("TxnAmt")).SendKeys("10000");
            //SelectElement dropdown = new SelectElement(driverIE.FindElement(By.Id("Mon")));
            

            driverIE.FindElement(By.Name("btnOK")).SendKeys(Keys.Enter);

            //Thread.Sleep(3000);
            /*acts.SendKeys(Keys.Enter);
            acts.Build().Perform();*/

            driverIE.FindElement(By.Name("rblTipo")).Click();
            driverIE.FindElement(By.Id("txtLogin")).SendKeys("BFPJUAMOC");
            driverIE.FindElement(By.Id("txtPassword")).SendKeys("PRUEBAS");

            driverIE.FindElement(By.Id("btnAuthOK")).SendKeys(Keys.Enter);

            /*Thread.Sleep(3000);
            acts.SendKeys(Keys.Enter);
            acts.SendKeys(Keys.Enter);
            acts.SendKeys(Keys.Enter);
            acts.Build().Perform();*/

            //System.Windows.Automation.AutomationElement t2 = WinFormAdapter.GetDialogBox("Sistema de Ventanillas WebTeller Banco Financiero - Windows Internet Explorer");

            Thread.Sleep(3000);
            Process[] p = Process.GetProcessesByName("iexplore");
            //Process p = Process.GetProcessById(22076);
            System.Windows.Automation.AutomationElement t1 = null;

            for (int i = 0; i <= p.Length; i++)
            {
                //System.Windows.Automation.AutomationElement t = WinFormAdapter.GetAEFromHandle(p[i].MainWindowHandle);
                System.Windows.Automation.AutomationElement t = AutomationElement.FromHandle(p[i].MainWindowHandle);
                t1 = WinFormAdapter.GetAEOnChildByName(t, "Impresora Apagada o Desconectada");

                if (t1 != null)
                {
                    break;
                }
            }
            //WinFormAdapter.ClickElement(WinFormAdapter.GetAEOnDescByName(t1, "Sí"));

            System.Windows.Automation.AutomationElement x = t1.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Sí"));
            //System.Windows.Automation.AutomationElementCollection x1 = t1.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Sí"));


            if (x != null)
            {
                var invokePattern = x.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                invokePattern.Invoke();
            }

            /*for (int i = 0; i <= _0_Descendants_1.Count; i++)
            {
                var invokePattern = _0_Descendants_1[i].GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                invokePattern.Invoke();
            }*/



            /*System.Windows.Point point = _0_Descendants_1[0].GetClickablePoint();
            Mouse.Move((int)point.X, (int)point.Y);
            Mouse.Click(MouseButton.Left);*/

            Thread.Sleep(3000);


            //driverIE.FindElement(By.Id("btnAlertaDismiss_1506115069882")).Click();
            //driverIE.FindElement(By.ClassName("botonApagado")).Click();
            //driverIE.FindElement(By.CssSelector("id^=btnAlertaDismiss")).Click();
            driverIE.FindElement(By.XPath("//*[contains(@id,'btnAlertaDismiss')]")).Click();
            driverIE.FindElement(By.XPath("//*[contains(@id,'btnAlertaDismiss')]")).Click();

            //acts.SendKeys(Keys.Enter);
            //acts.SendKeys(Keys.Enter);
        }

    }
}
