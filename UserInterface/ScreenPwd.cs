using MainframeAutomationSample.Entities;
using System.Diagnostics;
using TCSRPA;
using System.Reflection;
using IMS;
using MainframeAutomationSample.HelperClasses;
using System;
using System.Threading;

namespace MainframeAutomationSample.UserInterface
{
    public class ScreenPwd : IUserInterface
    {
        private EntityPwd objEntity = new EntityPwd();
        private Stopwatch objWatch = new Stopwatch();

        TCSRPA.XMLLibrary.ScreenIdentifier incorrectUserIdText;
        string incorrectUserIdMsg;

        public void DoActivities()
        {
            EntityGlobal.lastScreenName = EntityPwd.constScrName;

            CheckCompleteLoad();
            ExtractElements();
            DoUIValidation();
            DoOperations();
        }

        public void CheckCompleteLoad()
        {
            SystemLog.LogAuditMessage(EntityPwd.constScrName, DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_START, EntityGlobal.transactionId);

            TCSRPA.XMLLibrary.ScreenIdentifier validatePwdScreen = XMLLibrary.GetScreenIdentifier(EntityPwd.constScrName, EntityPwd.constValidatePwdScreen);

            string pwdScreenText = null;

            objWatch.Restart();
            while (pwdScreenText == null)
            {
                if (objWatch.ElapsedMilliseconds < EntityGlobal.pageLoadWaitTime)
                {
                    pwdScreenText = MainFrameAdapter.GetScreenText(validatePwdScreen.Row, validatePwdScreen.Column, validatePwdScreen.Length);
                    if (!(pwdScreenText.Equals(validatePwdScreen.Text)))
                    {
                        throw new UINotSupported(EntityPwd.constScrName);
                    }
                }
                else
                {
                    throw new UINotSupported(EntityPwd.constScrName);
                }
            }

            SystemLog.LogAuditMessage(EntityPwd.constScrName, DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_END, EntityGlobal.transactionId);
        }

        public void ExtractElements()
        {
            incorrectUserIdText = XMLLibrary.GetScreenIdentifier(EntityPwd.constScrName, EntityPwd.constIncorrectUserIdLbl);
            incorrectUserIdMsg = MainFrameAdapter.GetScreenText(incorrectUserIdText.Row, incorrectUserIdText.Column, incorrectUserIdText.Length);
        }

        public void DoUIValidation()
        {
            if (!(String.IsNullOrWhiteSpace(incorrectUserIdMsg)) && (incorrectUserIdMsg.Equals(incorrectUserIdText.Text)))
            {
                throw new AppException("Incorrect User ID");
            }
        }

        public void DoOperations()
        {
            string ispfText = XMLLibrary.GetName(EntityPwd.constScrName, EntityPwd.constIspf);
            TCSRPA.XMLLibrary.TextSetter pwdValue = XMLLibrary.GetTextSetter(EntityPwd.constScrName, EntityPwd.constPwdValue);

            MainFrameAdapter.SetTextOnScreen(pwdValue.Row, pwdValue.Column, CryptSecurity.ConvertToUnsecureString(EntityGlobal.pswd));
            Thread.Sleep(1000);
            MainFrameAdapter.SendKey(PcomKeys.Enter);

            #region Navigate till Ispf screen

            //ICH70001I RPASAR   LAST ACCESS
            MainFrameAdapter.SendKey(PcomKeys.Enter);

            //Notice screen
            MainFrameAdapter.SendKey(PcomKeys.Enter);

            //Notice screen - continue
            MainFrameAdapter.SendKey(PcomKeys.Enter);

            //ZIOEC screen 
            MainFrameAdapter.SendKey(PcomKeys.Enter);
           
            MainFrameAdapter.SendKey(ispfText);
            Thread.Sleep(1000);
            MainFrameAdapter.SendKey(PcomKeys.Enter);

            #endregion

            SystemLog.LogAuditMessage(EntityPwd.constScrName, DiagnosticLevel.level2, MethodBase.GetCurrentMethod().Name + EntityGlobal.METHOD_END, EntityGlobal.transactionId);
        }
    }
}

