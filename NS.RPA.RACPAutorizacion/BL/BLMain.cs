using EHLLAPI;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using To.AtNinjas.Util;

namespace NS.RPA.RACPAutorizacion
{
    class BLMain
    {
        CustomEHLLAPI ehllapi;
        private const int ROW_LENGTH = 80;
        public BLMain()
        {
            ehllapi = new CustomEHLLAPI();
            //ehllapi.Connect(SessionId);
        }

        private List<RelationshipOwed> RelationshipOweds { get; set; }
        private List<ProductChange> ProductChanges { get; set; }

        public void GetRelationshipOwedDataFromExcelConfig(string excelPath)
        {
            List<RelationshipOwed> lstRelationshipOwedData = null;
            try
            {
                using (var package = new ExcelPackage(new FileInfo(excelPath)))
                {
                    var workbook = package.Workbook;

                    var worksheet = workbook.Worksheets[1];
                    var totalRow = worksheet.Dimension.End.Row;

                    //string patron = @"[^\d+]";
                    //Regex regex = new Regex(patron);

                    lstRelationshipOwedData = new List<RelationshipOwed>();

                    for (int i = 2; i <= totalRow; i++)
                    {
                        //if (worksheet.Cells["B" + i].Value != null && worksheet.Cells["C" + i].Value != null &&
                        //    worksheet.Cells["D" + i].Value != null && worksheet.Cells["E" + i].Value != null &&
                        //    worksheet.Cells["F" + i].Value != null)
                        if (!String.IsNullOrEmpty(Convert.ToString(worksheet.Cells["A" + i].Value).Trim()) &&
                            !String.IsNullOrEmpty(Convert.ToString(worksheet.Cells["B" + i].Value).Trim()))
                        {
                            if (Convert.ToString(worksheet.Cells["C" + i].Value).Trim() == "CARGADO")
                            {
                                RelationshipOwed relationshipOwedData = new RelationshipOwed()
                                {
                                    ContractNumberDue = (worksheet.Cells["B" + i].Value ?? "").ToString().Trim(),
                                    FundingNumber = (worksheet.Cells["A" + i].Value ?? "").ToString().Trim()
                                };

                                lstRelationshipOwedData.Add(relationshipOwedData);
                            }
                        }
                        else
                        {
                            RelationshipOwed relationshipOwedData = new RelationshipOwed()
                            {
                                State = "ERROR"
                            };
                            lstRelationshipOwedData.Add(relationshipOwedData);
                        }
                    }
                }
            }
#pragma warning disable CS0168 // La variable 'e' se ha declarado pero nunca se usa
            catch (Exception e)
#pragma warning restore CS0168 // La variable 'e' se ha declarado pero nunca se usa
            {
                //TODO: especificar error que no se ha leído correctamente el archivo de configuracón disbursement config
                //Methods.LogProceso(e.ToString());
            }
            RelationshipOweds = lstRelationshipOwedData;
        }


        public void GetProductChangeDataFromExcelConfig(string excelPath)
        {
            List<RelationshipOwed> lstRelationshipOwedData = null;
            try
            {
                using (var package = new ExcelPackage(new FileInfo(excelPath)))
                {
                    var workbook = package.Workbook;

                    var worksheet = workbook.Worksheets[1];
                    var totalRow = worksheet.Dimension.End.Row;

                    //string patron = @"[^\d+]";
                    //Regex regex = new Regex(patron);

                    lstRelationshipOwedData = new List<RelationshipOwed>();

                    for (int i = 2; i <= totalRow; i++)
                    {
                        //if (worksheet.Cells["B" + i].Value != null && worksheet.Cells["C" + i].Value != null &&
                        //    worksheet.Cells["D" + i].Value != null && worksheet.Cells["E" + i].Value != null &&
                        //    worksheet.Cells["F" + i].Value != null)
                        if (!String.IsNullOrEmpty(Convert.ToString(worksheet.Cells["A" + i].Value).Trim()) &&
                            !String.IsNullOrEmpty(Convert.ToString(worksheet.Cells["B" + i].Value).Trim()))
                        {
                            if (Convert.ToString(worksheet.Cells["C" + i].Value).Trim() == "CARGADO")
                            {
                                RelationshipOwed relationshipOwedData = new RelationshipOwed()
                                {
                                    ContractNumberDue = (worksheet.Cells["B" + i].Value ?? "").ToString().Trim(),
                                    FundingNumber = (worksheet.Cells["A" + i].Value ?? "").ToString().Trim()
                                };

                                lstRelationshipOwedData.Add(relationshipOwedData);
                            }
                        }
                        else
                        {
                            RelationshipOwed relationshipOwedData = new RelationshipOwed()
                            {
                                State = "ERROR"
                            };
                            lstRelationshipOwedData.Add(relationshipOwedData);
                        }
                    }
                }
            }
#pragma warning disable CS0168 // La variable 'e' se ha declarado pero nunca se usa
            catch (Exception e)
#pragma warning restore CS0168 // La variable 'e' se ha declarado pero nunca se usa
            {
                //TODO: especificar error que no se ha leído correctamente el archivo de configuracón disbursement config
                //Methods.LogProceso(e.ToString());
            }
            RelationshipOweds = lstRelationshipOwedData;
        }

        public List<T> GetDataFromScreenList<T>(Func<string, T> formatMethod, int startYCoordenates,
         int pageSize, int rowsByItem, Func<string, int, string> readData/*MyDelegateType<T> formatMethod, */)
        {
            //string temp = ehllapi.ReadScreen(startYCoordenates + ",1", ROW_LENGTH * rowsByItem);
            string temp = readData(startYCoordenates.ToString(), ROW_LENGTH * rowsByItem);

            EhllapiWrapper.Wait();

            List<T> rows = new List<T>();
            int iItems = 0;

            if (temp.Trim() != "")
            {
                var obj = formatMethod(temp);

                if (obj != null)
                {
                    rows.Add(obj);
                    iItems = rowsByItem;
                }
            }

            int i = rowsByItem;
            while (temp.Trim() != "")
            {
                if (i >= pageSize)
                {
                    //ehllapi.SendStr("@v");

                    //TODO: si hay error en IBS ahí queda( deberíamos ver como obtener el valor de la x)
                    //temp = ehllapi.ReadScreen(startYCoordenates + ",1", ROW_LENGTH * rowsByItem);
                    temp = readData(startYCoordenates.ToString(), ROW_LENGTH * rowsByItem);
                    EhllapiWrapper.Wait();

                    Methods.LogProceso("last " + temp);
                    Methods.LogProceso("last2 " + rows[rows.Count - iItems / rowsByItem].ToString());


                    if (temp == rows[rows.Count - iItems / rowsByItem].ToString()) { break; }
                    else
                    {
                        var obj = formatMethod(temp);

                        if (obj != null)
                        {
                            rows.Add(obj);
                        }
                    }

                    i = rowsByItem;
                    iItems = rowsByItem;
                }
                else
                {
                    //TODO: depende de donde inicia la lista
                    //temp = ehllapi.ReadScreen(startYCoordenates + i + ",1", ROW_LENGTH * rowsByItem);
                    temp = readData((startYCoordenates + i).ToString(), ROW_LENGTH * rowsByItem);
                    EhllapiWrapper.Wait();

                    if (temp.Trim() != "")
                    {
                        var obj = formatMethod(temp);

                        if (obj != null)
                        {
                            rows.Add(obj);
                            iItems += rowsByItem;
                        }
                    }
                    i += rowsByItem;
                }
            }

            return rows;
        }


        public List<T> SetDataFromScreenList<T>(Func<string, int, T> formatMethod, int startYCoordenates,
         int pageSize, int rowsByItem, Func<string, int, string> readData/*MyDelegateType<T> formatMethod, */)
        {
            //string temp = ehllapi.ReadScreen(startYCoordenates + ",1", ROW_LENGTH * rowsByItem);
            string temp = readData(startYCoordenates.ToString(), ROW_LENGTH * rowsByItem);

            EhllapiWrapper.Wait();

            List<T> rows = new List<T>();
            int iItems = 0;

            if (temp.Trim() != "")
            {
                var obj = formatMethod(temp, startYCoordenates);

                if (obj != null)
                {
                    rows.Add(obj);
                    iItems = rowsByItem;
                }
            }

            int i = rowsByItem;
            while (temp.Trim() != "")
            {
                if (i >= pageSize)
                {
                    ehllapi.SendStr("@v");

                    //TODO: si hay error en IBS ahí queda( deberíamos ver como obtener el valor de la x)
                    //temp = ehllapi.ReadScreen(startYCoordenates + ",1", ROW_LENGTH * rowsByItem);
                    temp = readData(startYCoordenates.ToString(), ROW_LENGTH * rowsByItem);
                    EhllapiWrapper.Wait();

                    Methods.LogProceso("last " + temp);
                    Methods.LogProceso("last2 " + rows[rows.Count - iItems / rowsByItem].ToString());


                    if (temp == rows[rows.Count - iItems / rowsByItem].ToString()) { break; }
                    else
                    {
                        var obj = formatMethod(temp, startYCoordenates);

                        if (obj != null)
                        {
                            rows.Add(obj);
                        }
                    }

                    i = rowsByItem;
                    iItems = rowsByItem;
                }
                else
                {
                    //TODO: depende de donde inicia la lista
                    //temp = ehllapi.ReadScreen(startYCoordenates + i + ",1", ROW_LENGTH * rowsByItem);
                    temp = readData((startYCoordenates + i).ToString(), ROW_LENGTH * rowsByItem);
                    EhllapiWrapper.Wait();

                    if (temp.Trim() != "")
                    {
                        var obj = formatMethod(temp, (startYCoordenates + i));

                        if (obj != null)
                        {
                            rows.Add(obj);
                            iItems += rowsByItem;
                        }
                    }
                    i += rowsByItem;
                }
            }

            return rows;
        }

        public void SetApprovalPass(string pass)
        {
            ehllapi.SetCursorPos("3,14");
            EhllapiWrapper.Wait();
            ehllapi.SendStr(pass);
            EhllapiWrapper.Wait();
        }

        public string ReadScreen(string position, int lenght)
        {
            return ehllapi.ReadScreen(position + ",1", lenght);
        }

        public RelationshipOwed GetRelationshipOwed(string plot, int yCoordenates)
        {
            //me aseguro de que no tendré un fuera de rango... .
            //plot = plot.PadRight(80 * rowsByItem);

            //TODO: manejo de nulos
            //if (plot.Substring(20, 10)

            //ContractNumberDue or Funding?
            RelationshipOwed e = RelationshipOweds.FirstOrDefault(x => plot.Substring(3, 12).Trim() == x.ContractNumberDue);

            if (e != null)
            {
                e.State = "AUTORIZADO";
                ehllapi.SetCursorPos(yCoordenates + ",2");
                EhllapiWrapper.Wait();
                ehllapi.SendStr("I");
                EhllapiWrapper.Wait();

                ehllapi.SendStr("@E");
                EhllapiWrapper.Wait();
                ehllapi.SendStr("@E");
                EhllapiWrapper.Wait();
                ehllapi.SendStr("@E");
                EhllapiWrapper.Wait();
                ehllapi.SendStr("@E");
                EhllapiWrapper.Wait();
                ehllapi.SendStr("@E");
                EhllapiWrapper.Wait();
                ehllapi.SendStr("@E");
                EhllapiWrapper.Wait();
                ehllapi.SendStr("@E");
                EhllapiWrapper.Wait();

                ehllapi.SetCursorPos(yCoordenates + ",2");
                EhllapiWrapper.Wait();
                ehllapi.SendStr("Y");
                EhllapiWrapper.Wait();
                ehllapi.SendStr("@E");
                EhllapiWrapper.Wait();
            }

            e.Plot = plot;
            return e;
        } 

        public ProductChange GetProductChange(string plot, int yCoordenates)
        {
            //ContractNumberDue or Funding?
            ProductChange e = ProductChanges.FirstOrDefault(x => plot.Substring(3, 12).Trim() == x.FundingNumber);

            if (e != null)
            {
                e.State = "AUTORIZADO";
                ehllapi.SetCursorPos(yCoordenates + ",2");
                EhllapiWrapper.Wait();
                ehllapi.SendStr("X");
                EhllapiWrapper.Wait();

                ehllapi.SendStr("@E");
                EhllapiWrapper.Wait();

                ehllapi.SendStr("@1");
                EhllapiWrapper.Wait();
                ehllapi.SendStr("@E");
                EhllapiWrapper.Wait();
            }

            e.Plot = plot;
            return e;
        }

        public static decimal? DecimalTryParse(string value)
        {
            decimal result;
            if (!decimal.TryParse(value, out result))
                return null;
            return result;
        }

        public static DateTime? DatetimeTryParse(string value)
        {
            DateTime result;
            if (!DateTime.TryParse(value, out result))
                return null;
            return result;
        }

    }
}
