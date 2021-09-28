using EHLLAPI;
using MAFER.FinanciamientoDelExterior.Enum;
using MAFER.FinanciamientoDelExterior.ListConfig;
using OfficeOpenXml;
using System;
using System.IO;
using System.Text.RegularExpressions;
using To.AtNinjas.Util;

namespace MAFER.FinanciamientoDelExterior.Proccess
{
    public class IBSAutorizaPrestamoExterior
    {

        private string Password;
        public string SessionId;
        public int _Id;
        public string PasswordCreation;
        public string TemplateLoad;
        CustomEHLLAPI ehllapi;
        public ConfigModelProduct config;
        //  public ExcelManagament excel;


        public IBSAutorizaPrestamoExterior(String SessionId, string Passwd, int _Id)
        {
            this.Password = Passwd;
            ehllapi = new CustomEHLLAPI();
            ehllapi.Connect(SessionId);
            this._Id = _Id;
        }

        public void PrestamoAutorizacion(string MFTemplateAprobacion)
        {

            Methods.LogProceso("Menú principal de IBS..");

            ehllapi.SendStr("@7");
            Methods.GetMenuProgram("12");

            FileInfo existingFile = new FileInfo(MFTemplateAprobacion);

            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                int rowLoad = 0;
                string NroFinanciamiento = "", SiZeCustomer, TypeProduct = "";

                ExcelWorksheet worksheetAutoriza_load = package.Workbook.Worksheets[2];

                int rowCount = worksheetAutoriza_load.Dimension.End.Row;

                Methods.LogProceso("INICIA PROCESO DE DESCARGA DE FINANCIAMIENTOS" + "\n " + "con lecctura de excel.... " + TemplateLoad);

                string patron = @"[^\d+]";
                Regex regex = new Regex(patron);

                for (rowLoad = 2; rowLoad <= rowCount; rowLoad++)
                {
                    try
                    {
                        NroFinanciamiento = worksheetAutoriza_load.Cells["A" + rowLoad].Value == null ? "" : worksheetAutoriza_load.Cells["A" + rowLoad].Value.ToString().Trim();
                        NroFinanciamiento = regex.Replace(NroFinanciamiento, "");

                        if (NroFinanciamiento.Equals(""))
                        {
                            break;
                        }

                        SiZeCustomer = worksheetAutoriza_load.Cells["B" + rowLoad].Value == null ? "" : worksheetAutoriza_load.Cells["B" + rowLoad].Value.ToString().Trim();
                        TypeProduct = worksheetAutoriza_load.Cells["C" + rowLoad].Value == null ? "" : worksheetAutoriza_load.Cells["C" + rowLoad].Value.ToString().Trim();

                        Console.WriteLine("+++++++++ NroFinanciamiento+++:: " + NroFinanciamiento);
                        Console.WriteLine("+++++++++ ZiseCustomer+++:: " + SiZeCustomer);
                        Console.WriteLine("+++++++++ TypeProduct+++:: " + TypeProduct);

                        ehllapi.SetCursorPos("9,51");
                        ehllapi.SendStr("@F");
                        ehllapi.SendStr(NroFinanciamiento.Trim());
                        //ehllapi.SendStr("@A@+");
                        //ehllapi.SendStr("@E");
                        //EhllapiWrapper.Wait();

                        ehllapi.SetCursorPos("15,51");
                        ehllapi.SendStr("@F");
                        ehllapi.SendStr("01");

                        // SubProducto
                        var emp = "";

                        ExcelManagament excel = new ExcelManagament();
                        config = excel.GetDataConfig(MFTemplateAprobacion);
                        config.ListConfigCatalog.ForEach(delegate (ByProduct pr)
                        {
                            if (SiZeCustomer.Equals(pr.SizeCustomerType.ToString().Trim()))
                            {
                                if (TypeProduct.Equals(EnumTipoProducto._FEXP))
                                {
                                    emp = pr.Product_Type_Fexp.ToString().Trim();
                                }
                                else if (TypeProduct.Equals(EnumTipoProducto._FIMP))
                                {
                                    emp = pr.Product_Type_Fimp.ToString().Trim();
                                }
                            }
                        });

                        if (!emp.Equals(""))
                        {
                            worksheetAutoriza_load.Cells["D" + rowLoad].Value = emp;

                            ehllapi.SetCursorPos("17,51");
                            ehllapi.SendStr("@F");
                            ehllapi.SendStr(emp);

                            ehllapi.SendStr("@E"); // 
                            EhllapiWrapper.Wait();

                            worksheetAutoriza_load.Cells["E" + rowLoad].Value = "CARGADO";
                        }
                        else
                        {
                            worksheetAutoriza_load.Cells["E" + rowLoad].Value = "NO CARGADO_ " + "Código de Sub Producto incorrecto";
                        }

                        ehllapi.SendStr("@1"); // F1
                        EhllapiWrapper.Wait();

#if DEBUG
                        ehllapi.SendStr("@E"); // Intro
                        EhllapiWrapper.Wait();

#endif
                        ///**************
                        //      ehllapi.SendStr("@7"); // F7
                        //      EhllapiWrapper.Wait();
                        ///**************

                        NroFinanciamiento = ""; SiZeCustomer = ""; emp = "";

                        package.Save();

                    }
                    catch (Exception ex)
                    {
                        Methods.LogProceso("ERROR: " + ex.Message);
                    }
                }
            } // END for1
            Methods.LogProceso("FINALIZÓ EL PROCESO, presionar INTRO para salir...");
        } // End Using



    }
}
