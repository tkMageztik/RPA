using MainframeAutomationSample.HelperClasses;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using TCSRPA;

namespace MainframeAutomationSample.UserInterface
{
    public class ScreenBajaCtaUsuario : IUserInterface
    {
        public void DoActivities()
        {
            EntityGlobal.lastScreenName = "ScreenBajaCtaUsuario";

            CheckCompleteLoad();
            // ExtractElements();
            // DoUIValidation();
            DoOperations();
        }

        public void CheckCompleteLoad()
        {

        }

        public void DoOperations()
        {
            SystemLog.LogAuditMessage("ScreenBajaCtaUsuario", DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_START, EntityGlobal.transactionId);

            MainFrameAdapter.SetTextOnScreen(11, 53, "ROBOPRU2");

            Thread.Sleep(1000);
            MainFrameAdapter.SendKey(PcomKeys.Enter);

            SystemLog.LogAuditMessage("ScreenBajaCtaUsuario", DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_END, EntityGlobal.transactionId);
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
