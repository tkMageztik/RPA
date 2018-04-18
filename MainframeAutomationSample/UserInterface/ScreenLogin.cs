using IMS;
using MainframeAutomationSample.Entities;
using MainframeAutomationSample.HelperClasses;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Security;
using TCSRPA;

namespace MainframeAutomationSample.UserInterface
{
    class ScreenLogin : IUserInterface
    {
        private Stopwatch objWatch = new Stopwatch();

        public void DoActivities()
        {
            EntityGlobal.lastScreenName = "Login";

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



            MainFrameAdapter.WaitForCursor(6, 53);
            //GlobalHelper.GetUserPwd(EntityGlobal.appName);
            SystemLog.LogAuditMessage(EntityUserId.constScrName, DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_START, EntityGlobal.transactionId);

            EntityGlobal.userId = ConvertToSecureString("BFPGENJUL");
            EntityGlobal.pswd = ConvertToSecureString("JUP1T3R6");

            //TCSRPA.XMLLibrary.TextSetter userIdValue = XMLLibrary.GetTextSetter(EntityUserId.constScrName, EntityUserId.constUserIdValue);
            MainFrameAdapter.SetTextOnScreen(6, 53, CryptSecurity.ConvertToUnsecureString(EntityGlobal.userId));
            MainFrameAdapter.SetTextOnScreen(7, 53, CryptSecurity.ConvertToUnsecureString(EntityGlobal.pswd));

            Thread.Sleep(1000);
            MainFrameAdapter.SendKey(PcomKeys.Enter);

            SystemLog.LogAuditMessage(EntityUserId.constScrName, DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_END, EntityGlobal.transactionId);

        }

        public void DoUIValidation()
        {
            throw new NotImplementedException();
        }

        public void ExtractElements()
        {
            throw new NotImplementedException();
        }

        private SecureString ConvertToSecureString(string password)
        {
            if (password == null)
                throw new ArgumentNullException("password");

            var securePassword = new SecureString();

            foreach (char c in password)
                securePassword.AppendChar(c);

            securePassword.MakeReadOnly();
            return securePassword;
        }
    }
}
