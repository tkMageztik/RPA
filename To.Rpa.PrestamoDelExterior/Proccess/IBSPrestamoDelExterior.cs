using EHLLAPI;
using OfficeOpenXml;
using System;
using System.IO;
using System.Text.RegularExpressions;
using To.AtNinjas.Util;

namespace To.Rpa.PrestamoDelExterior.Proccess
{
    public class IBSPrestamoDelExterior
    {
        private string Password;
        public string SessionId;
        public int _Id;
        public string PasswordCreation;
        public string TemplateLoad;
        CustomEHLLAPI ehllapi;


        public IBSPrestamoDelExterior(String SessionId, string Passwd, int _Id)
        {
            this.Password = Passwd;
            ehllapi = new CustomEHLLAPI();
            ehllapi.Connect(SessionId);
            this._Id = _Id;
        }

        public void PrestamoExteriorLoad(string TemplateLoad, string Paswrord_2)
        {

#if DEBUG
            ehllapi.SendStr("@g");
            ehllapi.SetCursorPos("20,7");
            ehllapi.SendStr("GO MGCIA2");
            ehllapi.SendStr("@E");
#endif
            Methods.LogProceso("Menú principal de IBS..");
            Methods.GetMenuProgram("88");
            Methods.GetMenuProgram("4");
            Methods.GetMenuProgram("2");

            FileInfo existingFile = new FileInfo(TemplateLoad);

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                int rowLoad = 0;
                string NroFinanciamiento = "", NroContratoAdeudado = "", msgErrorAdeudado = "";

                ExcelWorksheet worksheet_load = package.Workbook.Worksheets[1];
                int rowCount = worksheet_load.Dimension.End.Row;

                Methods.LogProceso("INICIA PROCESO DE DESCARGA DE FINANCIAMIENTOS" + "\n " + "con lecctura de excel.... " + TemplateLoad);

                string patron = @"[^\d+]";
                Regex regex = new Regex(patron);

                for (rowLoad = 2; rowLoad <= rowCount; rowLoad++)
                {
                    try
                    {
                        NroFinanciamiento = worksheet_load.Cells["A" + rowLoad].Value == null ? "" : worksheet_load.Cells["A" + rowLoad].Value.ToString().Trim();
                        NroFinanciamiento = regex.Replace(NroFinanciamiento, "");

                        if (NroFinanciamiento.Equals(""))
                        {
                            break;
                        }

                        NroContratoAdeudado = worksheet_load.Cells["B" + rowLoad].Value == null ? "" : worksheet_load.Cells["B" + rowLoad].Value.ToString().Trim();
                        NroContratoAdeudado = regex.Replace(NroContratoAdeudado, "");

                        Console.WriteLine("+++++++++ NroFinanciamiento+++:: " + NroFinanciamiento);
                        Console.WriteLine("+++++++++ NroContratoAdeudado+++:: " + NroContratoAdeudado);

                        ehllapi.SetCursorPos("15,47");

#if DEBUG
                        //***** Para Desarrollo**************************************

                        var subProducto = "";
                        subProducto = ehllapi.ReadScreen("6,47", 4).Trim();

                        if (subProducto.Equals(""))
                        {
                            ehllapi.SetCursorPos("6,47");
                            ehllapi.SendStr("@a");
                            EhllapiWrapper.Wait();

                            ehllapi.SetCursorPos("12,25");
                            ehllapi.SendStr("FEXP");
                            for (int i = 0; i < 11; i++)
                            {
                                ehllapi.SendStr("@v");
                                EhllapiWrapper.Wait();
                            }

                            ehllapi.SetCursorPos("15,71");
                            ehllapi.SendStr("X");
                            ehllapi.SendStr("@E");
                            EhllapiWrapper.Wait();
                            //***********************************************************
                        }
#endif
                        ehllapi.SetCursorPos("15,47");
                        ehllapi.SendStr("@F");
                        ehllapi.SendStr(NroFinanciamiento.Trim());
                        ehllapi.SendStr("@A@+");
                        //ehllapi.SendStr("@E");
                        EhllapiWrapper.Wait();

                        ehllapi.SetCursorPos("17,47");
                        ehllapi.SendStr(Paswrord_2.Trim());
                        ehllapi.SendStr("@E");
                        EhllapiWrapper.Wait();

                        // Nro de Contrato Adeudado
                        ehllapi.SetCursorPos("20,68");
                        ehllapi.SendStr("@F");
                        ehllapi.SendStr(NroContratoAdeudado);

                        ehllapi.SendStr("@E"); // Intro
                        EhllapiWrapper.Wait();

                        msgErrorAdeudado = ehllapi.ReadScreen("5,7", 15).Trim();
                        if (msgErrorAdeudado.Trim().Equals("Dato no Valido"))
                        {
                            worksheet_load.Cells["C" + rowLoad].Value = "ERROR_" + msgErrorAdeudado.Trim();
                            ehllapi.SendStr("@E"); // Intro
                            EhllapiWrapper.Wait();
                        }
                        else
                        {
                            worksheet_load.Cells["C" + rowLoad].Value = "CARGADO";
                        }

                        ehllapi.SendStr("@1"); // F1
                        EhllapiWrapper.Wait();

                        //**************
                        //ehllapi.SendStr("@7"); // F7
                        //EhllapiWrapper.Wait();
                        //**************

#if DEBUG
                        // Esto se debe quitar es de prueba
                        ehllapi.SendStr("@1"); // F1
                        EhllapiWrapper.Wait();
#endif
                        ehllapi.SendStr("@6"); // F6
                        EhllapiWrapper.Wait();

                        ehllapi.SendStr("@3"); //F3
                        EhllapiWrapper.Wait();

                        NroFinanciamiento = ""; NroContratoAdeudado = ""; msgErrorAdeudado = "";
                        package.Save();

                    }
                    catch (Exception ex)
                    {
                        Methods.LogProceso("ERROR: " + ex.Message);
                    }
                }
            } // END for1

            Methods.LogProceso("FINALIZÓ 1ER PROCESO DE CARGA DE FINANCIAMIENTOS.");
            Methods.LogProceso("Presionar INTRO para continuar con 2do PROCESO...");
        } // End Using




    }
}
