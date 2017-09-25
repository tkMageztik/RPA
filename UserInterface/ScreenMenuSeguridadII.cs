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
    public class ScreenMenuSeguridadII : IUserInterface
    {
        public void DoActivities()
        {
            EntityGlobal.lastScreenName = "ScreenMenuSeguridadII";

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
            SystemLog.LogAuditMessage("ScreenMenuSeguridadII", DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_START, EntityGlobal.transactionId);

            MainFrameAdapter.SetTextOnScreen(22, 7, "1");
            

            Thread.Sleep(1000);
            MainFrameAdapter.SendKey(PcomKeys.Enter);

            SystemLog.LogAuditMessage("ScreenMenuSeguridadII", DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_END, EntityGlobal.transactionId);

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
