using System;
using TCSRPA;
using System.Threading;
using System.Reflection;
using MainframeAutomationSample.HelperClasses;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Windows.Automation;

namespace MainframeAutomationSample
{
    static class MainframeAutomation
    {
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll")]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")]
        public static extern IntPtr PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, IntPtr dwExtraInfo);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        public static void sendKeystroke(ushort k)
        {
            const uint WM_KEYDOWN = 0x100;
            const uint WM_SYSCOMMAND = 0x018;
            const uint SC_CLOSE = 0x053;

            IntPtr WindowToFind = FindWindow(null, "Untitled1 - Notepad++");

            IntPtr result3 = SendMessage(WindowToFind, WM_KEYDOWN, ((IntPtr)k), (IntPtr)0);
            //IntPtr result3 = SendMessage(WindowToFind, WM_KEYUP, ((IntPtr)c), (IntPtr)0);
        }

        [STAThread]
        static void Main_2()
        {
            bool isOpened = false;
            try
            {
                //Initializing mandatory parameters for application
                ConfigLibrary.Initialize();

                SystemLog.LogAuditMessage(MethodBase.GetCurrentMethod().Name, DiagnosticLevel.level1, EntityGlobal.ROBO_START, typeof(MainframeAutomation).Name);

                Initialize();

                //isOpened = MainFrameAdapter.OpenPcomSessionAndConnect();
                //Process p = Process.GetProcesses().SingleOrDefault(item => item.ProcessName == "pcsws");
//                Process.GetProcessById("")
                Process p = Process.GetProcesses().SingleOrDefault(item => item.ProcessName == "filezilla");
                System.Windows.Automation.AutomationElement t = WinFormAdapter.GetAEFromHandle(p.MainWindowHandle);

                //System.Windows.Automation.ControlType
                System.Windows.Automation.AutomationElement fdsfasdf = WinFormAdapter.GetAEOnChildByName(t, "ID_QUICKCONNECTBAR");
                //System.Windows.Automation.AutomationElement fdsfasdf = WinFormAdapter.GetAEOnChildByName(t, "Servidor:");

                WinFormAdapter.SetTextInTextArea(t, "fdsafsafasfdasfdsafdasfasfasdfasdfdsafsdafdasfdasfasd " + "\r\n" + "gfsdfdafs" + "\t" + "gadsfdad");
                WinFormAdapter.SetTextInTextArea(t, "servidor" + "\t" + "usuario" + "\t" + "contraseña");

                /*System.Windows.Automation.AutomationElement Servidor = WinFormAdapter.GetAEOnChildById(fdsfasdf, "-31800");
                System.Windows.Automation.AutomationElement Servidor1 = WinFormAdapter.GetAEOnChildById(fdsfasdf, "31801");*/
                System.Windows.Automation.AutomationElement Servidor = WinFormAdapter.GetAEOnChildByName(fdsfasdf,ControlType.Edit, "Servidor:");
                System.Windows.Automation.AutomationElement conexionRapida = WinFormAdapter.GetAEOnChildById(fdsfasdf, "-31925");
                //System.Windows.Automation.AutomationElement conexionRapida = WinFormAdapter.GetAEOnChildByName(fdsfasdf, "Conexión rápida");
                WinFormAdapter.SetText(Servidor, "hola");

                //Point conexionRapida.GetClickablePoint();

                 WinFormAdapter.ClickElement(conexionRapida);

                //WinFormAdapter.SetText(WinFormAdapter.GetAEOnChildByName(fdsfasdf, "Conexión rápida"), "rrorro");

                //System.Windows.Automation.AutomationElement t2 = WinFormAdapter.GetDialogBox("");

                t.SetFocus();
                //WinFormAdapter.SetText(t, "eeee");
                WinFormAdapter.SetTextInTextArea(t, "fdsafsafasfdasfdsafdasfasfasdfasdfdsafsdafdasfdasfasd " + "\r\n" + "gfsdfdafs" + "\t" + "gadsfdad");
                string texto = WinFormAdapter.GetText(t); //>> devuelve:  //Microsoft Excel - Libro1

                t.GetRuntimeId();
                t.SetFocus();
                System.Windows.Point punto = new System.Windows.Point(40,40);
                

                System.Windows.Automation.AutomationElement xx = WinFormAdapter.GetAEFromBoundingRectangle(punto);

                
                string www = WinFormAdapter.GetText(xx);
                //WinFormAdapter.SetTextInTextArea(xx, "EFDSFSDAFSFDSAFASFA");
                IntPtr e = WinFormAdapter.GetHandleForAE(t);

                const uint WM_KEYDOWN = 0x100;
                const uint KEYEVENTF_KEYUP = 0x0002;
                /*IntPtr result3 = SendMessage(WindowToFind, WM_KEYDOWN, ((IntPtr)k), (IntPtr)0);*/
                SendMessage(e, WM_KEYDOWN, (IntPtr)Keys.Enter, (IntPtr)0);

                WinFormAdapter.SetWindowForeground(e);
                bool eeee = SetForegroundWindow(e);
                keybd_event((byte)((char)Keys.Enter), 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                keybd_event((byte)((char)Keys.Enter), 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                keybd_event((byte)((char)Keys.Enter), 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                keybd_event((byte)((char)Keys.A), 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                keybd_event((byte)((char)Keys.A), 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                keybd_event((byte)((char)Keys.A), 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                keybd_event((byte)((char)Keys.A), 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                keybd_event((byte)((char)Keys.A), 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                keybd_event((byte)((char)Keys.Enter), 0, KEYEVENTF_KEYUP, IntPtr.Zero);
                keybd_event((byte)((char)Keys.Enter), 0, KEYEVENTF_KEYUP, IntPtr.Zero);

                //SendKeys.SendWait("EE");
                //SendKeys.SendWait(Keys.Enter.ToString());
                /*SendKeys.Send("K");
                SendKeys.Send(Keys.Enter.ToString());
                SendKeys.Send(Keys.Enter.ToString());
                SendKeys.Send(Keys.Enter.ToString());
                SendKeys.Send(Keys.Enter.ToString());
                SendKeys.Send(Keys.Enter.ToString());
                SendKeys.Send(Keys.Enter.ToString());*/


                


                Console.Write(www);
                MessageBox.Show(www);

                //System.Windows.Automation.AutomationElement t3 = WinFormAdapter.GetAEOnChildByName(t, "");


                //WinFormAdapter.ClickElement(t2);

                /*System.Windows.Automation.AutomationElement t3 = WinFormAdapter.GetAE()

                WinFormAdapter.SetText(t2, "BFPJUARUI");*/
                //Thread.Sleep(1000000);

                /*if (isOpened)
                {
                    new BusinessLogic.BLIspfValidator().DoActivities();
                }**/
                Thread.Sleep(10000900);
            }
            catch (TcsRpaException e)
            {
                SystemLog.LogErrorMessage(e, typeof(MainframeAutomation).Name, MethodBase.GetCurrentMethod().Name);
            }
            catch (Exception e)
            {
                SystemLog.LogErrorMessage(e, typeof(MainframeAutomation).Name, MethodBase.GetCurrentMethod().Name);
            }
            finally
            {
                if (isOpened)
                {
                    MainFrameAdapter.ClosePCOM();
                }
            }

            SystemLog.LogAuditMessage(MethodBase.GetCurrentMethod().Name, DiagnosticLevel.level1, EntityGlobal.ROBO_END, typeof(MainframeAutomation).Name);
            Thread.Sleep(2000);
        }

        private static void Initialize()
        {
            EntityGlobal.appLoadWaitTime = Convert.ToInt32(CommonHelper.GetConfigValueByKey(EntityGlobal.APP_LOAD_WAIT_TIME));
            EntityGlobal.pageLoadWaitTime = Convert.ToInt32(CommonHelper.GetConfigValueByKey(EntityGlobal.PAGE_LOAD_WAIT_TIME));
            EntityGlobal.appName = CommonHelper.GetConfigValueByKey(EntityGlobal.APP_NAME);
        }


    }

    
}
