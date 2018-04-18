using MainframeAutomationSample.Entities;
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
    public class ScreenMenuSeguridad : IUserInterface
    {
        public void DoActivities()
        {
            EntityGlobal.lastScreenName = "ScreenMenuSeguridad";

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
            SystemLog.LogAuditMessage("ScreenMenuSeguridad", DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_START, EntityGlobal.transactionId);

            MainFrameAdapter.SetTextOnScreen(22, 7, "87");

            Thread.Sleep(1000);
            MainFrameAdapter.SendKey(PcomKeys.Enter);

            SystemLog.LogAuditMessage("ScreenMenuSeguridad", DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_END, EntityGlobal.transactionId);

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
