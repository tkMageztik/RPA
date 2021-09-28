using MAFER.FinanciamientoDelExterior.Proccess;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Security;
using System.Windows.Automation;
using To.AtNinjas.Util;
using To.Rpa.PrestamoDelExterior.Proccess;

namespace To.Rpa.PrestamoDelExterior
{
    public class AppPrestamoDelExterior
    {

        public static int processID;
        public static string MFEmulatorPath;
        public static string MFEmulatorProcessName;
        public static string SessionId;
        public static string Username;
        public static string Password;
        public static string Paswrord_2;
        public string EmulatorURL;
        public string PasswordCreation;
        public string MFTemplateLoad;
        public string MFTemplateAutoriza;
        public int _Id;
        public IBSPrestamoDelExterior ibs;

        private Dictionary<string, object> CurrentConfig { get; set; }

        public AppPrestamoDelExterior()
        {
           

        }

        private void Init() {

            //TODO: Para configurar máximos intentos
            UInt16 attempts = 0;

            CurrentConfig = new Dictionary<string, object>();

            attempts = Convert.ToUInt16(ConfigurationManager.AppSettings["Attempts"]);
            EmulatorURL = ConfigurationManager.AppSettings["MFEmulatorPath"].ToString().Trim();
            MFEmulatorProcessName = ConfigurationManager.AppSettings["MFEmulatorProcessName"].ToString().Trim();
            SessionId = ConfigurationManager.AppSettings["SessionId"].Trim().ToString();
            Username = ConfigurationManager.AppSettings["MFUser"].ToString().Trim();
            Password = ConfigurationManager.AppSettings["MFPass"].ToString().Trim();
            MFTemplateLoad = ConfigurationManager.AppSettings["MFTemplateLoad"];
            MFTemplateAutoriza = ConfigurationManager.AppSettings["MFTemplateAutoriza"];
            Paswrord_2 = ConfigurationManager.AppSettings["PasswordCreation"].ToString().Trim();
        }

        public void DoActivities()
        {
            Init();
            int Id;
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
                if (Methods.OpenWS(EmulatorURL, SessionId, CurrentConfig["User"].ToString(), Password, out Id) == true)
                {
                    IBSPrestamoDelExterior loadOp = new IBSPrestamoDelExterior(SessionId, Password, _Id);
                    loadOp.PrestamoExteriorLoad(MFTemplateLoad, Paswrord_2);

                    Console.ReadKey();

                    IBSAutorizaPrestamoExterior loadAutorizacion = new IBSAutorizaPrestamoExterior(SessionId, Password, _Id);
                    loadAutorizacion.PrestamoAutorizacion(MFTemplateLoad);
                }
                else { }
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

        #region ignore
        private void DoOperations(string v1, object user, string v2, string password, string v3, string passwordCreation)
        {
            throw new NotImplementedException();
        }

        public IntPtr GetWindowsHandleByProcessId(int processID)
        {
            Process[] processes = Process.GetProcessesByName(MFEmulatorProcessName);

            for (int i = 0; i < processes.Length; i++)
            {
                if (processes[i].Id == processID)
                {
                    return processes[i].MainWindowHandle;
                }
            }
            return IntPtr.Zero;
        }


        public T Waiter<T>(Func<T, T> converter)
        {
            //return /* code to convert the setting to T... */
            //T t = (T)converter(Convert.ChangeType(1, typeof(T)));

            T t;
            do
            {
                t = (T)Convert.ChangeType(GetWindowsHandleByProcessId(1), typeof(T));

            } while (t.Equals((T)Convert.ChangeType(IntPtr.Zero, typeof(T))));

            return (T)Convert.ChangeType(t, typeof(T));
        }

        public void SomeUtility<T>(Func<object, T> converter)
        {
            var myType = converter("foo");
        }

        public static T ConfigSetting<T>(string settingName)
        {
            //return /* code to convert the setting to T... */
            return (T)Convert.ChangeType("", typeof(T));
        }

        public void CloseEmulator()
        {
            Process p = Process.GetProcessById(processID);

            if (p != null)
            {
                //Methods.Sleep();
                IntPtr t;
                do
                {
                    t = GetWindowsHandleByProcessId(p.Id);

                } while (t == IntPtr.Zero);

                AutomationElement _0 = AutomationElement.FromHandle(t);

                if (_0 != null)
                {
                    //AutomationElementCollection _0_Descendants_3 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Cancelar"));
                    AutomationElementCollection _0_Descendants_3;
                    do
                    {
                        //_0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Cancelar"));
                        _0_Descendants_3 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Cancelar"));

                    } while (_0_Descendants_3.Count == 0);

                    //AutomationElement _0 = AutomationElement.FromHandle(t);


                    var invokePattern = _0_Descendants_3[0].GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                    invokePattern.Invoke();
                    Methods.Sleep();

                }

                //Methods.Sleep();
                //bool tt = (bool)_0.GetCurrentPropertyValue(AutomationElement.IsOffscreenProperty);
                //SpinWait.SpinUntil(() => !tt, 5000);

                var windowPattern = _0.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
                windowPattern.Close();
            }
        }


        #endregion
    }
}
