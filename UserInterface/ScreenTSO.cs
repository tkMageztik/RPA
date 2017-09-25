using System.Diagnostics;
using System.Reflection;
using MainframeAutomationSample.Entities;
using TCSRPA;
using MainframeAutomationSample.HelperClasses;
using System.Threading;

namespace MainframeAutomationSample.UserInterface
{
    public class ScreenTSO : IUserInterface
    {
        private EntityTSO objEntity = new EntityTSO();
        private Stopwatch objStopWatch = new Stopwatch();

        public void DoActivities()
        {
            EntityGlobal.lastScreenName = EntityTSO.constScrName;

            CheckCompleteLoad();
            // ExtractElements();
            // DoUIValidation();
            DoOperations();
        }

        public void CheckCompleteLoad()
        {
            SystemLog.LogAuditMessage(EntityTSO.constScrName, DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_START, EntityGlobal.transactionId);

            TCSRPA.XMLLibrary.ScreenIdentifier applText = XMLLibrary.GetScreenIdentifier(EntityTSO.constScrName, EntityTSO.constAppText);
            string appText = null;

            objStopWatch.Restart();

            while (appText == null)
            {
                if (objStopWatch.ElapsedMilliseconds < EntityGlobal.pageLoadWaitTime)
                {
                    appText = MainFrameAdapter.GetScreenText(applText.Row, applText.Column, applText.Length);
                    if (!(appText.Equals(applText.Text)))
                    {
                        throw new UINotSupported(EntityTSO.constScrName);
                    }
                }
                else
                {
                    throw new UINotSupported(EntityTSO.constScrName);
                }
            }

            SystemLog.LogAuditMessage(EntityTSO.constScrName, DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_END, EntityGlobal.transactionId);
        }

        public void ExtractElements()
        {

        }

        public void DoUIValidation()
        {

        }

        public void DoOperations()
        {
            TCSRPA.XMLLibrary.TextSetter AppName = XMLLibrary.GetTextSetter(EntityTSO.constScrName, EntityTSO.constAppName);

            MainFrameAdapter.SetTextOnScreen(AppName.Row, AppName.Column, AppName.Text);
            Thread.Sleep(1000);
            MainFrameAdapter.SendKey(PcomKeys.Enter);

            SystemLog.LogAuditMessage(EntityTSO.constScrName, DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_END, EntityGlobal.transactionId);

        }
    }
}
