using MainframeAutomationSample.Entities;
using System.Diagnostics;
using TCSRPA;
using System.Reflection;
using MainframeAutomationSample.HelperClasses;
using System.Threading;

namespace MainframeAutomationSample.UserInterface
{
    public class ScreenIspf : IUserInterface
    {
        private EntityIspf objEntity = new EntityIspf();
        private Stopwatch objWatch = new Stopwatch();

        public void DoActivities()
        {
            EntityGlobal.lastScreenName = EntityIspf.constScrName;

            CheckCompleteLoad();
            // ExtractElements();
            //  DoUIValidation();
            DoOperations();
        }

        public void CheckCompleteLoad()
        {
            SystemLog.LogAuditMessage(EntityIspf.constScrName, DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_START, EntityGlobal.transactionId);


            TCSRPA.XMLLibrary.ScreenIdentifier menuText = XMLLibrary.GetScreenIdentifier(EntityIspf.constScrName, EntityIspf.constMenuText);

            string menuTextCheck = null;

            objWatch.Restart();
            while (menuTextCheck == null)
            {
                if (objWatch.ElapsedMilliseconds < EntityGlobal.pageLoadWaitTime)
                {

                    menuTextCheck = MainFrameAdapter.GetScreenText(menuText.Row, menuText.Column, menuText.Length);
                    if (!(menuTextCheck.Equals(menuText.Text)))
                    {
                        throw new UINotSupported(EntityIspf.constScrName);
                    }
                }
                else
                {
                    throw new UINotSupported(EntityIspf.constScrName);
                }
            }

            SystemLog.LogAuditMessage(EntityIspf.constScrName, DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_END, EntityGlobal.transactionId);
        }

        public void ExtractElements()
        {

        }

        public void DoUIValidation()
        {

        }

        public void DoOperations()
        {
            TCSRPA.XMLLibrary.TextSetter Logoff = XMLLibrary.GetTextSetter(EntityIspf.constScrName, EntityIspf.constLogoff);

            MainFrameAdapter.SendKey(PcomKeys.PF12);
            MainFrameAdapter.SetTextOnScreen(Logoff.Row, Logoff.Column, Logoff.Text);
            Thread.Sleep(1000);
            MainFrameAdapter.SendKey(PcomKeys.Enter);

            SystemLog.LogAuditMessage(EntityIspf.constScrName, DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_END, EntityGlobal.transactionId);
        }
    }
}
