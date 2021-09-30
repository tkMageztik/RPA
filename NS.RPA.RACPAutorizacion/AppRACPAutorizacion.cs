using EHLLAPI;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using To.AtNinjas.Util;

namespace NS.RPA.RACPAutorizacion
{
    public class AppRACPAutorizacion
    {
        public static int processID;
        public static string MFEmulatorPath;
        public static string MFEmulatorProcessName;
        public static string SessionId;
        public static string AuthorizationPass;
        //public static string Username;
        public static string Password;
        public string EmulatorURL;

        public static string MFTemplateLoad;
        private string StationID { get; set; }

        private Dictionary<string, object> CurrentConfig { get; set; }


        private BLMain _blMain;
        private BLScreenNavigation _blScreenNavigation;


        public void Init()
        {
            CurrentConfig = new Dictionary<string, object>();

            EmulatorURL = ConfigurationManager.AppSettings["MFEmulatorPath"].ToString().Trim();
            //MFEmulatorProcessName = ConfigurationManager.AppSettings["MFEmulatorProcessName"].ToString().Trim();
            SessionId = ConfigurationManager.AppSettings["SessionId"].Trim().ToString();
            //Username = ConfigurationManager.AppSettings["MFUser"].ToString().Trim();
            Password = ConfigurationManager.AppSettings["MFPass"].ToString().Trim();
            MFTemplateLoad = ConfigurationManager.AppSettings["MFTemplateLoad"];
            //Password = ConfigurationManager.AppSettings["Password"];
            //Paswrord_2 = ConfigurationManager.AppSettings["PasswordCreation"].ToString().Trim();
            AuthorizationPass = ConfigurationManager.AppSettings["AuthorizationPass"];

            _blScreenNavigation = new BLScreenNavigation();
            _blMain = new BLMain();

            _blMain.GetRelationshipOwedDataFromExcelConfig(MFTemplateLoad);


            _blMain.GetClientCodeDataFromExcelConfig(MFTemplateLoad);
            _blMain.GetProductChangeDataFromExcelConfig(MFTemplateLoad);

#if !DEBUG
            SetStationID();
#endif
        }

        private void SetStationID()
        {
            string text = File.ReadAllText(EmulatorURL);

            if (!StationID.Equals(""))
            {
                text = text.Replace("<StationID>", StationID);
                File.WriteAllText(EmulatorURL, text);
            }
        }

        public void DoActivities()
        {
            Init();
            try
            {
#if !DEBUG
                if (string.IsNullOrEmpty(Password))
                {
                    GetCredentials();
                }

                Password = new System.Net.NetworkCredential(string.Empty,
                    (SecureString)CurrentConfig["Password"]).Password;
#else
                CurrentConfig["User"] = "BFPROBOP2";
                Password = "BFPROBOP1";

#endif
                if (Methods.OpenWS(EmulatorURL, SessionId, CurrentConfig["User"].ToString(), Password, out int Id) == true)
                {
                    Base();
                }
                else { }

                Console.WriteLine("Fin");
                FlashWindow.FlashWindowEx(Process.GetCurrentProcess().MainWindowHandle);
                Console.ReadKey();
            }
            catch (Exception ex)
            {
                Methods.LogProceso("ERROR_2: " + ex.Message);
            }
        }

        private void GetCredentials()
        {
            Console.WriteLine("");
            Console.WriteLine("Por favor, ingrese su usuario de IBS");
            Console.WriteLine("");
            Console.WriteLine("Luego, presione ENTER para continuar ... .");

            CurrentConfig.Add("User", Console.ReadLine());

            Console.WriteLine("");
            Console.WriteLine("Por favor, ingrese el password de IBS");
            Console.WriteLine("Luego, presione ENTER para continuar ... .");

            Console.WriteLine("");
            Console.WriteLine("NOTA: Considerar que el password ingresado, se utilizará encriptado, " +
                "por lo que NO se almacenará y ni la aplicación conocerá su valor.");

            CurrentConfig.Add("Password", GetPassword());
            Console.WriteLine("");
        }

        public SecureString GetPassword()
        {
            var pwd = new SecureString();
            while (true)
            {
                ConsoleKeyInfo i = Console.ReadKey(true);
                if (i.Key == ConsoleKey.Enter)
                {
                    break;
                }
                else if (i.Key == ConsoleKey.Backspace)
                {
                    if (pwd.Length > 0)
                    {
                        pwd.RemoveAt(pwd.Length - 1);
                        Console.Write("\b \b");
                    }
                }
                else if (i.KeyChar != '\u0000') // KeyChar == '\u0000' if the key pressed does not correspond to a printable character, e.g. F1, Pause-Break, etc
                {
                    pwd.AppendChar(i.KeyChar);
                    Console.Write("*");
                }
            }
            return pwd;
        }

        public void Base()
        {
            //for ??

            if (_blScreenNavigation.ShowScreenRATransactionsApproval())
            {
                _blMain.SetApprovalPass(AuthorizationPass);

                var relationshipOweds = _blMain.SetDataFromScreenList(_blMain.GetRelationshipOwed, 7, 14, 1, _blMain.ReadScreen);

                _blScreenNavigation.BackToMainMenu();                
            }

            if (_blScreenNavigation.ShowScreenClientPosition())
            {
                _blMain.UpdateCompanySize();
                _blScreenNavigation.BackToMainMenu2();
            }

            if (_blScreenNavigation.ShowScreenCPApproval())
            {
                var productChanges = _blMain.SetDataFromScreenList(_blMain.GetProductChange, 7, 14, 1, _blMain.ReadScreen);
              
                _blScreenNavigation.BackToMainMenu();
            }

        }

    }
}
