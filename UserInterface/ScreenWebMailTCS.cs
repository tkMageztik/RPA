using MainframeAutomationSample.HelperClasses;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Automation;
using TCSRPA;

namespace MainframeAutomationSample.UserInterface
{
    class ScreenWebMailTCS : IUserInterface
    {

        public void CheckCompleteLoad()
        {
        }

        public void DoActivities()
        {
            EntityGlobal.lastScreenName = "ScreenWebMailTCS";
            CheckCompleteLoad();
            // ExtractElements();
            // DoUIValidation();
            DoOperations();
        }

        public void DoOperations()
        {

            /*
             * 
            Process[] p = Process.GetProcessesByName("chrome");

            System.Windows.Automation.AutomationElement _0 = AutomationElement.FromHandle(p[0].MainWindowHandle);

            System.Windows.Automation.AutomationElementCollection _0_Descendants_1 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "ID de usuario:"));

            */

            /*
            IWebDriver driverChrome = new ChromeDriver(@"D:\COMPARTIDO_PUBLICO\RPA\TCS RPA V1.2 - Training\Sample\Automation - Proceso Cese - JJRDC\Mainframe Automation\Mainframe Automation Solution\packages\Selenium.WebDriver.ChromeDriver.2.32.0\driver\win32");
            
            driverChrome.Navigate().GoToUrl("https://webmail.tcs.com/");
            driverChrome.FindElement(By.Name("Username")).SendKeys("juan.ruizdecastilla@tcs.com");

            driverChrome.FindElement(By.Name("Password")).SendKeys("Santiago230611@16");
            driverChrome.FindElement(By.Id("mybutton")).SendKeys(Keys.Enter);
            */

            /*
            InternetExplorerOptions ieOptions = new InternetExplorerOptions();
            ieOptions.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
            
            IWebDriver driverIE = new InternetExplorerDriver(@"D:\COMPARTIDO_PUBLICO\RPA\TCS RPA V1.2 - Training\Sample\Automation - Proceso Cese - JJRDC\Mainframe Automation\Mainframe Automation Solution\packages\Selenium.WebDriver.IEDriver.3.5.1\driver", ieOptions);
            
            driverIE.Navigate().GoToUrl("https://webmail.tcs.com/");
            driverIE.FindElement(By.Name("Username")).SendKeys("juan.ruizdecastilla@tcs.com");
            
            driverIE.FindElement(By.Name("Password")).SendKeys("Santiago230611@16");
            driverIE.FindElement(By.Id("mybutton")).SendKeys(Keys.Enter);
            */

            InternetExplorerOptions ieOptions = new InternetExplorerOptions();
            ieOptions.IntroduceInstabilityByIgnoringProtectedModeSettings = true;

            //IWebDriver driverIE = new InternetExplorerDriver(@"D:\COMPARTIDO_PUBLICO\RPA\TCS RPA V1.2 - Training\Sample\Automation - Proceso Cese - JJRDC\Mainframe Automation\Mainframe Automation Solution\packages\Selenium.WebDriver.IEDriver.3.5.1\driver", ieOptions);

            IWebDriver driverIE = new InternetExplorerDriver(System.Configuration.ConfigurationManager.AppSettings["IEDriverPath"], ieOptions);
            //driverIE.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(60);
            driverIE.Navigate().GoToUrl("https://webmail.tcs.com/");

            driverIE.FindElement(By.Name("Username")).SendKeys("juan.ruizdecastilla");

            /*esta porción es porque mi contraseña lleva '@' */
            Actions acts = new Actions(driverIE);
            acts.KeyDown(Keys.Control).KeyDown(Keys.LeftAlt);
            acts.SendKeys("q");
            acts.KeyUp(Keys.Control).KeyUp(Keys.LeftAlt);
            acts.Build().Perform();
            /**/

            driverIE.FindElement(By.Name("Username")).SendKeys("tcs.com");

            driverIE.FindElement(By.Name("Password")).SendKeys("Santiago230611");

            /*esta porción es porque mi contraseña lleva '@', con la forma anterior no sirve en el campo contraseña*/
            acts.KeyDown(Keys.Control).KeyDown(Keys.Alt);
            acts.SendKeys(Keys.NumberPad6);
            acts.SendKeys(Keys.NumberPad4);
            acts.KeyUp(Keys.Control).KeyUp(Keys.Alt);
            acts.Build().Perform();
            /**/

            driverIE.FindElement(By.Name("Password")).SendKeys("16");

            driverIE.FindElement(By.Id("mybutton")).SendKeys(Keys.Enter);


            //WaitUntilElementExists(driverIE, By.Id("s_MainFrame"));


            //driverIE.Url = "http://somedomain/url_that_delays_loading";
            //IWebElement myDynamicElement = driverIE.FindElement(By.Id("s_MainFrame"));


            driverIE.SwitchTo().Frame("s_MainFrame");

            IWebElement WebElement = driverIE.FindElements(By.XPath("//div[@id='vListe-listview-container-mail']/div"))[1];




            int count = driverIE.FindElements(By.XPath("//div[@id='tble-listview-container-mail']/div")).Count;

            //IList<IWebElement> AllCheckBoxes = WebElement.FindElements(By.XPath("//div"));

            //for (int i=0; i < 

            //WebElement.FindElements(By.Id("e-listview-container-mail-row-3"))[0]

            int totalCorreos = int.Parse(Regex.Match(driverIE.FindElement(By.Id("e-mailoutline-row-($Inbox)1-elem-OUTLINEELEM5")).Text, "\\d+").ToString());

            IWebElement trash = driverIE.FindElement(By.Id("e-actions-mailview-inbox-trash"));


            //IWebElement ee = WebElement.FindElements(By.Id("e-listview-container-mail-row-2"))[0];

            for (int i = 0; i <= totalCorreos; i++)
            {
                IWebElement ele = WebElement.FindElements(By.Id("e-listview-container-mail-row-" + (i + 1)))[0];

                if (ele.Text.Contains("Mail Router"))
                {

                    //ele.Click();
                    ele.SendKeys(Keys.Delete);
                    //Una vez borrado el elemento ya no se puede hacer keys.ArrowDown, se tiene que tomar el nuevo elemento
                    ele = driverIE.SwitchTo().ActiveElement();
                    ele.SendKeys(Keys.ArrowDown);
                    i--;
                    //trash.Click();
                    Thread.Sleep(1000);

                }
                else
                {
                    ele.SendKeys(Keys.ArrowDown);
                }





            }

        }
        /*
         static void WaitForLoad(this IWebDriver driver, int timeoutSec = 15)
        {
            WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, timeoutSec));
            
            wait.Until(wd => ((IJavaScriptExecutor)wd).ExecuteScript("return document.readyState") == "complete");
        }*/

        //this will search for the element until a timeout is reached
        /*public static IWebElement WaitUntilElementExists(IWebDriver driver, By elementLocator, int timeout = 10)
        {
            try
            {
                IWebElement t;
                using (driver)
                {
                    )
                    var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(timeout));
                    //return wait.Until( ExpectedConditions.ElementExists(elementLocator));

                    t =  wait.Until<IWebElement>(d => d.FindElement(elementLocator));
                }
                //return t;


            }
            catch (NoSuchElementException)
            {
                Console.WriteLine("Element with locator: '" + elementLocator + "' was not found in current context page.");
                throw;
            }
        }*/

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
