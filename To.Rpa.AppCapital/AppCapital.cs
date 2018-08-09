using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;
using To.Rpa.AppCapital.BL;
using To.AtNinjas.Util;

namespace To.Rpa.AppCapital
{
    public class AppCapital
    {
        public static int processID;
        public static string MFEmulatorPath;
        public static string MFEmulatorProcessName;

        public AppCapital()
        {
            //TODO: Para configurar máximos intentos
            int i = 0;
            UInt16 attempts = 0;

            do
            {
                try
                {
                    attempts = Convert.ToUInt16(ConfigurationManager.AppSettings["Attempts"]);
                    MFEmulatorPath = ConfigurationManager.AppSettings["MFEmulatorPath"].ToString().Trim();
                    MFEmulatorProcessName = ConfigurationManager.AppSettings["MFEmulatorProcessName"].ToString().Trim();

                    Assembly assembly = Assembly.GetExecutingAssembly();
                    FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);

                    //TODO: La configuración puede estar a nivel de la torre de control
                    
                    Methods.LogProceso("Se inicia proceso de AppCapital: " + fvi.ProductName + " -> " + MethodBase.GetCurrentMethod().Name);
                    Methods.Sleep();
                    Process px = Process.Start(MFEmulatorPath);
                    processID = px.Id;
                    new BLMain().DoActivities();
                    Methods.LogProceso("Proceso finalizado");
                    Console.WriteLine("Presione ENTER para cerrar esta ventana.");
                    Console.ReadLine();
                    break;
                }
                catch (Exception exc)
                {
                    i++;
                    Methods.LogProceso(" ERROR: " + exc.StackTrace + " " + MethodBase.GetCurrentMethod().Name);
                    CloseEmulator();
                }
                finally
                {
                    Methods.LogProceso(" FINALLY: Intento nro: " + i);
                }
            }
            while (i < attempts);

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

        //public Object<> Waiter()
        //{
        //    IntPtr t;
        //    do
        //    {
        //        t = GetWindowsHandleByProcessId(p.Id);

        //    } while (t == IntPtr.Zero);

        //    return t;
        //}

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
    }
}
