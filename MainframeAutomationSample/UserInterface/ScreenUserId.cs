using System.Diagnostics;
using TCSRPA;
using MainframeAutomationSample.Entities;
using System.Reflection;
using IMS;
using MainframeAutomationSample.HelperClasses;
using System.Threading;

namespace MainframeAutomationSample.UserInterface
{
    public class ScreenUserId : IUserInterface
    {
        private EntityUserId objEntity = new EntityUserId();
        private Stopwatch objWatch = new Stopwatch();

        public void DoActivities()
        {
            EntityGlobal.lastScreenName = EntityUserId.constScrName;

            CheckCompleteLoad();
            // ExtractElements();
            // DoUIValidation();
            DoOperations();
        }

        public void CheckCompleteLoad()
        {
            SystemLog.LogAuditMessage(EntityUserId.constScrName, DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_START, EntityGlobal.transactionId);

            TCSRPA.XMLLibrary.ScreenIdentifier userIdLabel = XMLLibrary.GetScreenIdentifier(EntityUserId.constScrName, EntityUserId.constUserIdLabel);

            string userIdLabelCheck = null;

            objWatch.Restart();
            while (userIdLabelCheck == null)
            {
                if (objWatch.ElapsedMilliseconds < EntityGlobal.pageLoadWaitTime)
                {
                    userIdLabelCheck = MainFrameAdapter.GetScreenText(userIdLabel.Row, userIdLabel.Column, userIdLabel.Length);

                    if (!(userIdLabelCheck.Equals(userIdLabel.Text)))
                    {
                        throw new UINotSupported(EntityUserId.constScrName);
                    }
                }
                else
                {
                    throw new UINotSupported(EntityTSO.constScrName);
                }
            }

            SystemLog.LogAuditMessage(EntityUserId.constScrName, DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_END, EntityGlobal.transactionId);
        }

        public void ExtractElements()
        {

        }

        public void DoUIValidation()
        {

        }

        public void DoOperations()
        {
            GlobalHelper.GetUserPwd(EntityGlobal.appName);

            TCSRPA.XMLLibrary.TextSetter userIdValue = XMLLibrary.GetTextSetter(EntityUserId.constScrName, EntityUserId.constUserIdValue);
            MainFrameAdapter.SetTextOnScreen(userIdValue.Row, userIdValue.Column, CryptSecurity.ConvertToUnsecureString(EntityGlobal.userId));
            Thread.Sleep(1000);
            MainFrameAdapter.SendKey(PcomKeys.Enter);

            SystemLog.LogAuditMessage(EntityUserId.constScrName, DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_END, EntityGlobal.transactionId);

        }
    }
}
