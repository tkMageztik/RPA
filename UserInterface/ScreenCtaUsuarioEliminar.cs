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
    public class ScreenCtaUsuarioEliminar : IUserInterface
    {

        public void DoActivities()
        {
            EntityGlobal.lastScreenName = "ScreenCtaUsuarioEliminar";

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
            SystemLog.LogAuditMessage("ScreenCtaUsuarioEliminar", DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_START, EntityGlobal.transactionId);

            MainFrameAdapter.SetTextOnScreen(18, 51, "S");

            Thread.Sleep(1000);
            MainFrameAdapter.SendKey(PcomKeys.Enter);

            SystemLog.LogAuditMessage("ScreenCtaUsuarioEliminar", DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_END, EntityGlobal.transactionId);
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
