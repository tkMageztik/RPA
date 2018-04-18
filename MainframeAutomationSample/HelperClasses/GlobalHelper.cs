using IMS;
using System;

namespace MainframeAutomationSample.HelperClasses
{
    class GlobalHelper
    {
        public static void GetUserPwd(string appName)
        {
            string strFilePath = @"D:\Ravina\Mainframe Automation\Mainframe Automation\Mainframe Automation Solution\bin\Release";
            CryptSecurity.Decryption(appName, strFilePath, out EntityGlobal.userId, out EntityGlobal.pswd);
        }
    }

    public class AppException : Exception
    {
        public AppException()
        {

        }

        public AppException(string message)
            : base(message)
        {

        }
    }


}
