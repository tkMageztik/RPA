using System.Security;

namespace MainframeAutomationSample.HelperClasses
{
    class EntityGlobal
    {
        //constants
        public const string METHOD_START = " START";
        public const string METHOD_END = " END";
        public const string ROBO_START = " ROBO START";
        public const string ROBO_END = " ROBO END";
        public const string TRANSACTION_START = "TRANSACTION START";
        public const string TRANSACTION_END = "TRANSACTION END";


        //RoboConfig.xml parameters
        public const string APP_LOAD_WAIT_TIME = "AppLoadWaitTime";
        public const string PAGE_LOAD_WAIT_TIME = "PageWaitTime";
        public const string APP_NAME = "AppName";


        //variables
        public static int appLoadWaitTime;
        public static int pageLoadWaitTime;
        public static string lastScreenName;
        public static string transactionId;
        public static SecureString userId;
        public static SecureString pswd;
        public static string appName;

        public static bool isTransactionComplete = false;
    }
}
