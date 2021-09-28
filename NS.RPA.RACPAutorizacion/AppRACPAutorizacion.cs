using EHLLAPI;
using NS.RPA.RACPAutorizacion.BL;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
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
        public static string Username;
        public static string Password;
        public string EmulatorURL;

        public static string MFTemplateLoad;

        private Dictionary<string, object> CurrentConfig { get; set; }


        private BLMain _blMain;
        private BLScreenNavigation _blScreenNavigation;

        public void Init()
        {
            CurrentConfig = new Dictionary<string, object>();

            EmulatorURL = ConfigurationManager.AppSettings["MFEmulatorPath"].ToString().Trim();
            //MFEmulatorProcessName = ConfigurationManager.AppSettings["MFEmulatorProcessName"].ToString().Trim();
            SessionId = ConfigurationManager.AppSettings["SessionId"].Trim().ToString();
            Username = ConfigurationManager.AppSettings["MFUser"].ToString().Trim();
            Password = ConfigurationManager.AppSettings["MFPass"].ToString().Trim();
            MFTemplateLoad = ConfigurationManager.AppSettings["MFTemplateLoad"];
            //Paswrord_2 = ConfigurationManager.AppSettings["PasswordCreation"].ToString().Trim();

            _blScreenNavigation = new BLScreenNavigation();
            _blMain = new BLMain();
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
                    Test2();
                    //IBSPrestamoDelExterior loadOp = new IBSPrestamoDelExterior(SessionId, Password, _Id);
                    //loadOp.PrestamoExteriorLoad(MFTemplateLoad, Paswrord_2);

                    //Console.ReadKey();

                    //IBSAutorizaPrestamoExterior loadAutorizacion = new IBSAutorizaPrestamoExterior(SessionId, Password, _Id);
                    //loadAutorizacion.PrestamoAutorizacion(MFTemplateLoad);
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

        public void Test()
        {

            FileInfo existingFile = new FileInfo(MFTemplateLoad);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //int rowLoad = 0;
                string nroFinanciamiento = "", nroContratoAdeudado = "", msgErrorAdeudado = "";

                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                int rowCount = worksheet.Dimension.End.Row;

                //                Methods.LogProceso("INICIA PROCESO DE DESCARGA DE FINANCIAMIENTOS" + "\n " + "con lecctura de excel.... " + TemplateLoad);

                //                string patron = @"[^\d+]";
                //                Regex regex = new Regex(patron);

                //                for (rowLoad = 2; rowLoad <= rowCount; rowLoad++)
                //                {
                //                    try
                //                    {
                //                        NroFinanciamiento = worksheet_load.Cells["A" + rowLoad].Value == null ? "" : worksheet_load.Cells["A" + rowLoad].Value.ToString().Trim();
                //                        NroFinanciamiento = regex.Replace(NroFinanciamiento, "");

                //                        if (NroFinanciamiento.Equals(""))
                //                        {
                //                            break;
                //                        }

                //                        NroContratoAdeudado = worksheet_load.Cells["B" + rowLoad].Value == null ? "" : worksheet_load.Cells["B" + rowLoad].Value.ToString().Trim();
                //                        NroContratoAdeudado = regex.Replace(NroContratoAdeudado, "");

                //                        Console.WriteLine("+++++++++ NroFinanciamiento+++:: " + NroFinanciamiento);
                //                        Console.WriteLine("+++++++++ NroContratoAdeudado+++:: " + NroContratoAdeudado);

                //                        ehllapi.SetCursorPos("15,47");

                //#if DEBUG
                //                        //***** Para Desarrollo**************************************

                //                        var subProducto = "";
                //                        subProducto = ehllapi.ReadScreen("6,47", 4).Trim();

                //                        if (subProducto.Equals(""))
                //                        {
                //                            ehllapi.SetCursorPos("6,47");
                //                            ehllapi.SendStr("@a");
                //                            EhllapiWrapper.Wait();

                //                            ehllapi.SetCursorPos("12,25");
                //                            ehllapi.SendStr("FEXP");
                //                            for (int i = 0; i < 11; i++)
                //                            {
                //                                ehllapi.SendStr("@v");
                //                                EhllapiWrapper.Wait();
                //                            }

                //                            ehllapi.SetCursorPos("15,71");
                //                            ehllapi.SendStr("X");
                //                            ehllapi.SendStr("@E");
                //                            EhllapiWrapper.Wait();
                //                            //***********************************************************
                //                        }
                //#endif
                //                        ehllapi.SetCursorPos("15,47");
                //                        ehllapi.SendStr("@F");
                //                        ehllapi.SendStr(NroFinanciamiento.Trim());
                //                        ehllapi.SendStr("@A@+");
                //                        //ehllapi.SendStr("@E");
                //                        EhllapiWrapper.Wait();

                //                        ehllapi.SetCursorPos("17,47");
                //                        ehllapi.SendStr(Paswrord_2.Trim());
                //                        ehllapi.SendStr("@E");
                //                        EhllapiWrapper.Wait();

                //                        // Nro de Contrato Adeudado
                //                        ehllapi.SetCursorPos("20,68");
                //                        ehllapi.SendStr("@F");
                //                        ehllapi.SendStr(NroContratoAdeudado);

                //                        ehllapi.SendStr("@E"); // Intro
                //                        EhllapiWrapper.Wait();

                //                        msgErrorAdeudado = ehllapi.ReadScreen("5,7", 15).Trim();
                //                        if (msgErrorAdeudado.Trim().Equals("Dato no Valido"))
                //                        {
                //                            worksheet_load.Cells["C" + rowLoad].Value = "ERROR_" + msgErrorAdeudado.Trim();
                //                            ehllapi.SendStr("@E"); // Intro
                //                            EhllapiWrapper.Wait();
                //                        }
                //                        else
                //                        {
                //                            worksheet_load.Cells["C" + rowLoad].Value = "CARGADO";
                //                        }

                //                        ehllapi.SendStr("@1"); // F1
                //                        EhllapiWrapper.Wait();

                //                        //**************
                //                        //ehllapi.SendStr("@7"); // F7
                //                        //EhllapiWrapper.Wait();
                //                        //**************

                //#if DEBUG
                //                        // Esto se debe quitar es de prueba
                //                        ehllapi.SendStr("@1"); // F1
                //                        EhllapiWrapper.Wait();
                //#endif
                //                        ehllapi.SendStr("@6"); // F6
                //                        EhllapiWrapper.Wait();

                //                        ehllapi.SendStr("@3"); //F3
                //                        EhllapiWrapper.Wait();

                //                        NroFinanciamiento = ""; NroContratoAdeudado = ""; msgErrorAdeudado = "";
                //                        package.Save();

                //                    }
                //                    catch (Exception ex)
                //                    {
                //                        Methods.LogProceso("ERROR: " + ex.Message);
                //                    }

            } // END for1

        }

        public void Test2()
        {
            //for ??

            if (_blScreenNavigation.ShowScreenRATransactionsApproval(""))
            {
                _blMain.SetApprovalPass("8998");

                //_bLLoanPaymentMovement
                var loanPaymentMovements = _blMain.SetDataFromScreenList(_blMain.GetLoanPaymentMovement, 7, 12, 2, _blMain.ReadScreen);

                /*
                _blScreenNavigation.ShowScreenClientLoanPaymentMovements();

                loanPaymentMovements = _blMain.SetDataFromScreenList(_bLLoanPaymentMovement.GetLoanPaymentMovement, 7, 12, 2, _blMain.ReadScreen);

                _blMain.UpdateToNegativeAmount(loanPaymentMovements);

                groupedLoanPaymentMovements = _bLLoanPaymentMovement.GetGroupedLoanPaymentMovements(loanPaymentMovements);

                _blLoanDebitBalance.Save(newLoan, groupedLoanPaymentMovements);

                */
                _blScreenNavigation.BackToMainMenu();

                //loanPaymentMovements = _blMain.GetDataFromScreenList(_bLLoanPaymentMovement.GetLoanPaymentMovement, 7, 12, 2, _blMain.ReadScreen);
            }


            if (_blScreenNavigation.ShowScreenCPApproval(""))
            {
                //_bLLoanPaymentMovement
                var loanPaymentMovements = _blMain.SetDataFromScreenList(_blMain.GetLoanPaymentMovement, 7, 12, 2, _blMain.ReadScreen);

                /*
                _blScreenNavigation.ShowScreenClientLoanPaymentMovements();

                loanPaymentMovements = _blMain.SetDataFromScreenList(_bLLoanPaymentMovement.GetLoanPaymentMovement, 7, 12, 2, _blMain.ReadScreen);

                _blMain.UpdateToNegativeAmount(loanPaymentMovements);

                groupedLoanPaymentMovements = _bLLoanPaymentMovement.GetGroupedLoanPaymentMovements(loanPaymentMovements);

                _blLoanDebitBalance.Save(newLoan, groupedLoanPaymentMovements);

                */
                _blScreenNavigation.BackToMainMenu();

                //loanPaymentMovements = _blMain.GetDataFromScreenList(_bLLoanPaymentMovement.GetLoanPaymentMovement, 7, 12, 2, _blMain.ReadScreen);
            }

        }

    }
}
