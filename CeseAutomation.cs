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
    static class CeseAutomation
    {
        [STAThread]
        static void Main()
        {
            bool isOpened = false;
            try
            {
                //Initializing mandatory parameters for application
                ConfigLibrary.Initialize();

                SystemLog.LogAuditMessage(MethodBase.GetCurrentMethod().Name, DiagnosticLevel.level1, EntityGlobal.ROBO_START, typeof(MainframeAutomation).Name);

                Initialize();

                /*V1*/
                
                //isOpened = MainFrameAdapter.OpenPcomSessionAndConnect();
                
                //Process[] p = Process.GetProcessesByName("pcsws");

                /*
                System.Windows.Automation.AutomationElement t = WinFormAdapter.GetAEFromHandle(p[0].MainWindowHandle);

                System.Windows.Automation.AutomationElement t1 = WinFormAdapter.GetAEOnChildByName(t, "Inicio de sesión de IBM i");

                
                System.Windows.Automation.AutomationElement t2 = WinFormAdapter.GetAEOnChildByName(t, ControlType.Edit, "ID de usuario:");
                System.Windows.Automation.AutomationElement t2_1 = WinFormAdapter.GetAEOnChildById(t, ControlType.Edit,"1321");
                System.Windows.Automation.AutomationElement t2_2 = WinFormAdapter.GetAEOnDescByName(t, ControlType.Edit, "ID de usuario:");
                System.Windows.Automation.AutomationElement t2_3 = WinFormAdapter.GetAEOnDescByName(t, ControlType.Edit, "1321");
                
                System.Windows.Automation.AutomationElement t3 = WinFormAdapter.GetAEOnChildByName(t, ControlType.Edit, "Contraseña:");

                WinFormAdapter.SetText(t2, "BFPJUARUI");
                WinFormAdapter.SetText(t3, "BFPJUARUI2");
                */
                
                
                /*if (isOpened)
                {
                   new BusinessLogic.BLIspfValidator().DoActivities();
                }*/
                
                new BusinessLogic.BLIsWebValidator().DoActivities();
                //new BusinessLogic.BLIsDesktopValidator().DoActivities();
            }
            catch (TcsRpaException e)
            {
                SystemLog.LogErrorMessage(e, typeof(CeseAutomation).Name, MethodBase.GetCurrentMethod().Name);
            }
            catch (Exception e)
            {
                SystemLog.LogErrorMessage(e, typeof(CeseAutomation).Name, MethodBase.GetCurrentMethod().Name);
            }
            finally
            {
                if (isOpened)
                {
                    MainFrameAdapter.ClosePCOM();
                }
            }

            SystemLog.LogAuditMessage(MethodBase.GetCurrentMethod().Name, DiagnosticLevel.level1, EntityGlobal.ROBO_END, typeof(CeseAutomation).Name);
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
