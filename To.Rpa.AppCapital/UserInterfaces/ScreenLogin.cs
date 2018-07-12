using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using EHLLAPI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using S = DocumentFormat.OpenXml.Spreadsheet.Sheets;
using E = DocumentFormat.OpenXml.OpenXmlElement;
using A = DocumentFormat.OpenXml.OpenXmlAttribute;
using OfficeOpenXml;
using System.IO;

namespace To.Rpa.AppCapital.UserInterfaces
{
    public class ScreenLogin
    {
        //private Stopwatch objWatch = new Stopwatch();

        public void DoActivities()
        {
            //EntityGlobal.lastScreenName = "Login";

            CheckCompleteLoad();
            // ExtractElements();
            // DoUIValidation();
            DoOperations();

        }

        public void CheckCompleteLoad()
        {

        }

        public void DoOperations()
        {
            //TODO: Depende de que la sesión "A" esté libre

            var iConnection = EhllapiWrapper.Connect("A").ToString();

            Debug.WriteLine("Valor de conexión: " + iConnection);

            string lectura = null;
            int tiempo = 2000;
            if (iConnection == "0")
            {
                Thread.Sleep(tiempo);
                EhllapiWrapper.SetCursorPos(453);
                EhllapiWrapper.SendStr("BFPROBOP2");

                EhllapiWrapper.SetCursorPos(533);
                EhllapiWrapper.SendStr("BFPROBOP6");

                EhllapiWrapper.SendStr("@E");

                EhllapiWrapper.SetCursorPos(162);

                EhllapiWrapper.ReadScreen(162, 239, out lectura);

                //if (lectura.IndexOf("La cola de mensajes BFPJUARUI está asignada a otro trabajo") != 0)
                if (lectura.IndexOf("está asignada a otro trabajo") != -1)
                {
                    EhllapiWrapper.SendStr("@E");
                }

                EhllapiWrapper.SetCursorPos(1687);
                EhllapiWrapper.SendStr("88");
                EhllapiWrapper.SendStr("@E");
                EhllapiWrapper.SetCursorPos(1687);
                EhllapiWrapper.SendStr("4");
                EhllapiWrapper.SendStr("@E");
                EhllapiWrapper.SetCursorPos(1687);
                EhllapiWrapper.SendStr("2");
                EhllapiWrapper.SendStr("@E");

                //TODO: BORRAR
                //TODO: mover capital... .
                //string capital = "";
                //TODO: BORRAR

                //GetSheetInfo(@"D:\COMPARTIDO_PUBLICO\RPA\TCS RPA V1.2 - Training\Sample\RPA\Main\To.Rpa.AppCapital\bin\Debug\DATA_PROCESO.xlsx");

                //string t = GetCellValue(@"D:\COMPARTIDO_PUBLICO\RPA\TCS RPA V1.2 - Training\Sample\RPA\Main\To.Rpa.AppCapital\bin\Debug\DATA_PROCESO.xlsx",
                //    "DATA","A2");

                FileInfo existingFile = new FileInfo(@"d:\Users\juarui\Source\Repos\RPA\To.Rpa.AppCapital\DATA_PROCESO.xlsx");

                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                    int row = 2;
                    //TODO: aqui mover capital... .
                    string prestamo = "", banco, agencia, moneda, ctaTramite, cta, estado, capital, error;
                    decimal monto;

                    do
                    {
                        try
                        {
                            estado = worksheet.Cells[row, 8].Value.ToString();

                            if (estado.Equals("ERROR"))
                            {
                                continue;
                            }

                            prestamo = worksheet.Cells[row, 1].Value.ToString();
                            monto = Convert.ToDecimal(worksheet.Cells[row, 2].Value);
                            ctaTramite = worksheet.Cells[row, 6].Value == null ? "" : worksheet.Cells[row, 6].Value.ToString();
                            cta = worksheet.Cells[row, 7].Value == null ? "" : worksheet.Cells[row, 7].Value.ToString();

                            EhllapiWrapper.SetCursorPos(1167);
                            EhllapiWrapper.SendStr(prestamo);
                            EhllapiWrapper.SetCursorPos(1327);
                            EhllapiWrapper.SendStr("9999");
                            EhllapiWrapper.SendStr("@2");

                            EhllapiWrapper.SetCursorPos(244);
                            EhllapiWrapper.ReadScreen(244, 14, out capital);
                            capital = capital.Substring(0, 14);
                            //Convert.to strMonto
                            Thread.Sleep(tiempo);

                            //TODO: para que sirve esta validación?
                            if (capital != null)
                            {
                                Thread.Sleep(tiempo);

                                //TODO: validar en el excel que solo se puede ingresar un campo a la vez.
                                if (!ctaTramite.Equals(""))
                                {
                                    banco = worksheet.Cells[row, 3].Value.ToString();
                                    agencia = worksheet.Cells[row, 4].Value.ToString();
                                    moneda = worksheet.Cells[row, 5].Value.ToString();

                                    EhllapiWrapper.SetCursorPos(575);
                                    EhllapiWrapper.SendStr(monto.ToString());
                                    EhllapiWrapper.SendStr("@A@+");
                                    EhllapiWrapper.SendStr("@T");

                                    EhllapiWrapper.SetCursorPos(592);
                                    EhllapiWrapper.SendStr(banco);

                                    EhllapiWrapper.SetCursorPos(595);
                                    EhllapiWrapper.SendStr(agencia);

                                    EhllapiWrapper.SetCursorPos(599);
                                    EhllapiWrapper.SendStr(moneda);

                                    EhllapiWrapper.SetCursorPos(603);
                                    EhllapiWrapper.SendStr(ctaTramite);

                                }
                                else
                                {
                                    //TODO: CUENTA
                                }

                                

                                EhllapiWrapper.SetCursorPos(1682);
                                EhllapiWrapper.SendStr(prestamo + " / APLIC CAP");
                                EhllapiWrapper.SendStr("@b");
                                EhllapiWrapper.ReadScreen(244, 14, out error);

                                if (error == "mensaje de error")
                                {
                                    EhllapiWrapper.SendStr("@E");
                                    if ("" == "existen deducciones")
                                    {
                                        //TODO: UBICAR BIEN DEDUCCIONES
                                        EhllapiWrapper.SetCursorPos(1455);
                                        EhllapiWrapper.SendStr("@c");
                                    }
                                }

                                //TODO: CAPITAL
                                //MetodoRecursivo(monto);


                                EhllapiWrapper.SetCursorPos(899);
                                EhllapiWrapper.SendStr(".");
                                EhllapiWrapper.SendStr("@3");

#if DEBUG
                                EhllapiWrapper.SendStr("@3");
#endif

                                //PARA PASAR A OTRA OPERACIÓN
                                //EhllapiWrapper.SendStr("@7");
                                //EhllapiWrapper.SetCursorPos(1687);
                                //EhllapiWrapper.SendStr("2");
                                //EhllapiWrapper.SendStr("@E");
                                prestamo = worksheet.Cells[row + 1, 1].Value.ToString();

                            }
                        }
                        catch (Exception exc)
                        {
                            //TODO: guardar log.
                            worksheet.SetValue(row, 8, "ERROR");
                        }
                        finally { row++; }

                    } while (!prestamo.Equals(""));
                }

                //// Open the spreadsheet document for read-only access.
                //using (SpreadsheetDocument document =
                //SpreadsheetDocument.Open(@"d:\Users\juarui\Source\Repos\RPA\To.Rpa.AppCapital\DATA_PROCESO.xlsx", false))
                //{
                //    row1Col0 = GetCellValue(document, "DATA", "A2");
                //    row1Col1 = GetCellValue(document, "DATA", "B2");
                //    row1Col2 = GetCellValue(document, "DATA", "C2");
                //    row1Col3 = GetCellValue(document, "DATA", "D2");
                //    row1Col4 = GetCellValue(document, "DATA", "E2");
                //    row1Col5 = GetCellValue(document, "DATA", "F2");
                //    //row1Col5 = "2908070001210000";//dr[5].ToString();


                //    EhllapiWrapper.SetCursorPos(1167);
                //    EhllapiWrapper.SendStr(row1Col0);
                //    EhllapiWrapper.SetCursorPos(1327);
                //    EhllapiWrapper.SendStr("9999");
                //    EhllapiWrapper.SendStr("@2");

                //    EhllapiWrapper.SetCursorPos(244);
                //    EhllapiWrapper.ReadScreen(244, 14, out capital);

                //    Thread.Sleep(tiempo);

                //    if (capital != null)
                //    {
                //        Thread.Sleep(tiempo);

                //        EhllapiWrapper.SetCursorPos(575);
                //        EhllapiWrapper.SendStr(row1Col1);
                //        EhllapiWrapper.SendStr("@A@+");
                //        EhllapiWrapper.SendStr("@T");

                //        EhllapiWrapper.SetCursorPos(592);
                //        EhllapiWrapper.SendStr(row1Col2);

                //        EhllapiWrapper.SetCursorPos(595);
                //        EhllapiWrapper.SendStr(row1Col3);

                //        EhllapiWrapper.SetCursorPos(599);
                //        EhllapiWrapper.SendStr(row1Col4);

                //        EhllapiWrapper.SetCursorPos(603);
                //        EhllapiWrapper.SendStr(row1Col5);

                //        EhllapiWrapper.SetCursorPos(1682);
                //        //TODO:CONTADOR
                //        EhllapiWrapper.SendStr(row1Col0 + " / APLIC CAP  _" /*+ cont.ToString()*/);
                //        EhllapiWrapper.SendStr("@b");
                //        Thread.Sleep(tiempo);

                //        EhllapiWrapper.SetCursorPos(899);
                //        EhllapiWrapper.SendStr(".");
                //        EhllapiWrapper.SendStr("@3");

                //        //TODO: FALTA CONFIRMAR LA SALIDA CON F3 y el popup que sale... 


                //        //                    //string newColumn = "H";
                //        //                    //string newRow = "5";
                //        //                    //string worksheet2 = "DATA";

                //        //                    ////string sql2 = String.Format("UPDATE [{0}$] SET {1}{2}={3}", worksheet2, newColumn, newRow, "Hola");
                //        //                    //string commandString = String.Format("UPDATE [{0}${1}{2}:{1}{2}] SET F1='{3}'", worksheet2, newColumn, newRow, 1);
                //        //                    //OleDbCommand objCmdSelect = new OleDbCommand(commandString, connection);
                //        //                    //objCmdSelect.ExecuteNonQuery();



                //        //                    //COMANDOS PARA EL SIGUIENTE REGISTRO
                //        //                    //MainFrameAdapter.SendKey(PcomKeys.PF7);
                //        //                    //MainFrameAdapter.SetTextOnScreen(22, 7, "2");
                //        //                    //MainFrameAdapter.SendKey(PcomKeys.Enter);

                //        //PARA PASAR A OTRA OPERACIÓN
                //        EhllapiWrapper.SendStr("@7");
                //        EhllapiWrapper.SetCursorPos(1687);
                //        EhllapiWrapper.SendStr("2");
                //        EhllapiWrapper.SendStr("@E");
                //    }


                //}

            }
        }

        public void MetodoRecursivo(decimal monto)
        {
            decimal temp = 0;
            do
            {
                decimal interes = (monto - 0.01M) - 0;
                temp = interes;

            } while (0 < temp);
        }

        public static void GetSheetInfo(string fileName)
        {
            // Open file as read-only.
            using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(fileName, false))
            {
                S sheets = mySpreadsheet.WorkbookPart.Workbook.Sheets;

                // For each sheet, display the sheet information.
                foreach (E sheet in sheets)
                {
                    foreach (A attr in sheet.GetAttributes())
                    {
                        Debug.WriteLine("{0}: {1}", attr.LocalName, attr.Value);


                    }
                }
            }
        }

        // Retrieve the value of a cell, given a file name, sheet name, 
        // and address name.
        public static string GetCellValue(string fileName,
            string sheetName,
            string addressName)
        {
            string value = null;

            // Open the spreadsheet document for read-only access.
            using (SpreadsheetDocument document =
                SpreadsheetDocument.Open(fileName, false))
            {
                // Retrieve a reference to the workbook part.
                WorkbookPart wbPart = document.WorkbookPart;

                // Find the sheet with the supplied name, and then use that 
                // Sheet object to retrieve a reference to the first worksheet.
                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
                  Where(s => s.Name == sheetName).FirstOrDefault();

                // Throw an exception if there is no sheet.
                if (theSheet == null)
                {
                    throw new ArgumentException("sheetName");
                }

                // Retrieve a reference to the worksheet part.
                WorksheetPart wsPart =
                    (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                // Use its Worksheet property to get a reference to the cell 
                // whose address matches the address you supplied.
                Cell theCell = wsPart.Worksheet.Descendants<Cell>().
                  Where(c => c.CellReference == addressName).FirstOrDefault();

                // If the cell does not exist, return an empty string.
                if (theCell != null)
                {
                    value = theCell.InnerText;

                    // If the cell represents an integer number, you are done. 
                    // For dates, this code returns the serialized value that 
                    // represents the date. The code handles strings and 
                    // Booleans individually. For shared strings, the code 
                    // looks up the corresponding value in the shared string 
                    // table. For Booleans, the code converts the value into 
                    // the words TRUE or FALSE.
                    if (theCell.DataType != null)
                    {
                        switch (theCell.DataType.Value)
                        {
                            case CellValues.SharedString:

                                // For shared strings, look up the value in the
                                // shared strings table.
                                var stringTable =
                                    wbPart.GetPartsOfType<SharedStringTablePart>()
                                    .FirstOrDefault();

                                // If the shared string table is missing, something 
                                // is wrong. Return the index that is in
                                // the cell. Otherwise, look up the correct text in 
                                // the table.
                                if (stringTable != null)
                                {
                                    value =
                                        stringTable.SharedStringTable
                                        .ElementAt(int.Parse(value)).InnerText;
                                }
                                break;

                            case CellValues.Boolean:
                                switch (value)
                                {
                                    case "0":
                                        value = "FALSE";
                                        break;
                                    default:
                                        value = "TRUE";
                                        break;
                                }
                                break;
                        }
                    }
                }
            }
            return value;
        }

        public static string GetCellValue(SpreadsheetDocument document,
            string sheetName,
            string addressName)
        {
            string value = null;

            // Retrieve a reference to the workbook part.
            WorkbookPart wbPart = document.WorkbookPart;

            // Find the sheet with the supplied name, and then use that 
            // Sheet object to retrieve a reference to the first worksheet.
            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
              Where(s => s.Name == "DATA").FirstOrDefault();

            // Throw an exception if there is no sheet.
            if (theSheet == null)
            {
                throw new ArgumentException("sheetName");
            }

            // Retrieve a reference to the worksheet part.
            WorksheetPart wsPart =
                (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

            // Use its Worksheet property to get a reference to the cell 
            // whose address matches the address you supplied.
            Cell theCell = wsPart.Worksheet.Descendants<Cell>().
                  Where(c => c.CellReference == addressName).FirstOrDefault();

            // If the cell does not exist, return an empty string.
            if (theCell != null)
            {
                value = theCell.InnerText;

                // If the cell represents an integer number, you are done. 
                // For dates, this code returns the serialized value that 
                // represents the date. The code handles strings and 
                // Booleans individually. For shared strings, the code 
                // looks up the corresponding value in the shared string 
                // table. For Booleans, the code converts the value into 
                // the words TRUE or FALSE.
                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.SharedString:

                            // For shared strings, look up the value in the
                            // shared strings table.
                            var stringTable =
                                wbPart.GetPartsOfType<SharedStringTablePart>()
                                .FirstOrDefault();

                            // If the shared string table is missing, something 
                            // is wrong. Return the index that is in
                            // the cell. Otherwise, look up the correct text in 
                            // the table.
                            if (stringTable != null)
                            {
                                value =
                                    stringTable.SharedStringTable
                                    .ElementAt(int.Parse(value)).InnerText;
                            }
                            break;

                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                    }
                }
            }
            return value;
        }

        public void DoUIValidation()
        {
            throw new NotImplementedException();
        }

        public void ExtractElements()
        {
            throw new NotImplementedException();
        }

        private SecureString ConvertToSecureString(string password)
        {
            if (password == null)
                throw new ArgumentNullException("password");

            var securePassword = new SecureString();

            foreach (char c in password)
                securePassword.AppendChar(c);

            securePassword.MakeReadOnly();
            return securePassword;
        }
    }
}
