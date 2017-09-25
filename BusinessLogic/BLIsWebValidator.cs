﻿using MainframeAutomationSample.HelperClasses;
using MainframeAutomationSample.UserInterface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using TCSRPA;

namespace MainframeAutomationSample.BusinessLogic
{
    class BLIsWebValidator
    {
        int count_AppException = 0;
        int count_TcsRpaException = 0;
        int count_ComExceptiom = 0;
        int count_UnknownErrors = 0;

        public void DoActivities()
        {
            EntityGlobal.transactionId = "0001";

            //int businessTrails = 3;

            //for (int i = 0; i < businessTrails; i++)
            //{
            EntityGlobal.isTransactionComplete = false;

            SystemLog.LogAuditMessage(MethodBase.GetCurrentMethod().Name, DiagnosticLevel.level1, EntityGlobal.TRANSACTION_START, EntityGlobal.transactionId);
            //Thread.Sleep(EntityGlobal.pageLoadWaitTime);
            //MainFrameAdapter.WaitForCursor(3, 15);
            try
            {
                /*new ScreenTSO().DoActivities();
                new ScreenUserId().DoActivities();
                new ScreenPwd().DoActivities();
                new ScreenIspf().DoActivities();*/

                //new ScreenWebMailTCS().DoActivities();
                new ScreenWebTeller().DoActivities();

                EntityGlobal.isTransactionComplete = true;
                //break;
            }
            catch (AppException e)
            {
                count_AppException++;
                SystemLog.LogExceptionMessage(e.Message, EntityGlobal.transactionId, EntityGlobal.lastScreenName);
                if (count_AppException > 5)
                {
                    throw new TcsRpaException("Maximum limit reached for Application Exception");
                }

            }
            catch (TcsRpaException e)
            {
                count_TcsRpaException++;
                SystemLog.LogErrorMessage(e, EntityGlobal.transactionId, EntityGlobal.lastScreenName);
                if (count_TcsRpaException > 5)
                {
                    throw new TcsRpaException("Maximum limit reached for TCSRPA Exception");
                }
            }
            catch (COMException e)
            {
                count_ComExceptiom++;
                SystemLog.LogErrorMessage(e, EntityGlobal.transactionId, EntityGlobal.lastScreenName);
                if (count_ComExceptiom > 5)
                {
                    throw new TcsRpaException("Maximum limit reached for COM Exception");
                }
            }
            catch (Exception e)
            {
                count_UnknownErrors++;
                SystemLog.LogErrorMessage(e, EntityGlobal.transactionId, EntityGlobal.lastScreenName);
                if (count_UnknownErrors > 5)
                {
                    throw new TcsRpaException("Maximum limit reached for Unknown errors Exception");
                }
            }
            //}
            if (!EntityGlobal.isTransactionComplete)
            {
                SystemLog.LogExceptionMessage("Transaction Not Completed", EntityGlobal.transactionId, EntityGlobal.lastScreenName);
            }

            SystemLog.LogAuditMessage(MethodBase.GetCurrentMethod().Name, DiagnosticLevel.level1, EntityGlobal.TRANSACTION_END, EntityGlobal.transactionId);
        }
    }
}
