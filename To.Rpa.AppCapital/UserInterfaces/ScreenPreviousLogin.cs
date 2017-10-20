using DocumentFormat.OpenXml.Packaging;
using EHLLAPI;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using S = DocumentFormat.OpenXml.Spreadsheet.Sheets;
using E = DocumentFormat.OpenXml.OpenXmlElement;
using A = DocumentFormat.OpenXml.OpenXmlAttribute;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Tcs.Rpa.AppCapital.Mainframe.Interfaces
{
    public class ScreenPreviousLogin : IUserInterface
    {
        public void CheckCompleteLoad()
        {
        }

        public void DoActivities()
        {
            CheckCompleteLoad();
            // ExtractElements();
            // DoUIValidation();
            DoOperations();
        }

        public void DoOperations()
        {
            Process[] p = Process.GetProcessesByName("pcsws");
            //p[0].WaitForInputIdle();

            Thread.Sleep(2000);

            AutomationElement _0 = AutomationElement.FromHandle(p[0].MainWindowHandle);

            AutomationElementCollection _0_Descendants_1 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "ID de usuario:"));
            //WinFormAdapter.SetText(_0_Descendants_1[1], "BFPJUARUI");

            ValuePattern etb = _0_Descendants_1[1].GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
            etb.SetValue("BFPJUARUI");

            AutomationElementCollection _0_Descendants_2 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Contraseña:"));
            //WinFormAdapter.SetText(_0_Descendants_2[1], "BFPJUARUI2");

            ValuePattern etb2 = _0_Descendants_2[1].GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
            etb2.SetValue("BFPJUARUI3");

            //WinFormAdapter.ClickElement(WinFormAdapter.GetAEOnDescByName(_0, "Aceptar"));
            AutomationElementCollection _0_Descendants_3 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Aceptar"));

            var invokePattern = _0_Descendants_3[0].GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            invokePattern.Invoke();

            /************************** INTENTOS ******************/

            //intento 1
            //var tt = _0_Descendants_3[0].GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
            //invokePattern.Invoke();

            //intento 2
            //AutomationElementCollection xx = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Sesión A - [24 x 80]"));

            //AutomationElement w = xx[0].FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.AutomationIdProperty, "2"));
            //AutomationElement w1 = xx[0].FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.AutomationIdProperty, "2"));


            //AutomationProperty[] t = xx[0].GetSupportedProperties();

            /************************** INTENTOS ******************/



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
                EhllapiWrapper.SendStr("BFPROBOP4");

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


                string montoAbonado = null;
                string row1Col0 = "";
                string row1Col1 = "";
                string row1Col2 = "";
                string row1Col3 = "";
                string row1Col4 = "";
                string row1Col5 = "";

                //GetSheetInfo(@"D:\COMPARTIDO_PUBLICO\RPA\TCS RPA V1.2 - Training\Sample\RPA\Main\To.Rpa.AppCapital\bin\Debug\DATA_PROCESO.xlsx");

                //string t = GetCellValue(@"D:\COMPARTIDO_PUBLICO\RPA\TCS RPA V1.2 - Training\Sample\RPA\Main\To.Rpa.AppCapital\bin\Debug\DATA_PROCESO.xlsx",
                //    "DATA","A2");


                // Open the spreadsheet document for read-only access.
                using (SpreadsheetDocument document =
                    SpreadsheetDocument.Open(@"D:\COMPARTIDO_PUBLICO\RPA\TCS RPA V1.2 - Training\Sample\RPA\Main\To.Rpa.AppCapital\bin\Debug\DATA_PROCESO.xlsx", false))
                {
                    row1Col0 = GetCellValue(document, "DATA", "A2");
                    row1Col1 = GetCellValue(document, "DATA", "B2");
                    row1Col2 = GetCellValue(document, "DATA", "C2");
                    row1Col3 = GetCellValue(document, "DATA", "D2");
                    row1Col4 = GetCellValue(document, "DATA", "E2");
                    row1Col5 = GetCellValue(document, "DATA", "F2");
                    //row1Col5 = "2908070001210000";//dr[5].ToString();


                    EhllapiWrapper.SetCursorPos(1167);
                    EhllapiWrapper.SendStr(row1Col0);
                    EhllapiWrapper.SetCursorPos(1327);
                    EhllapiWrapper.SendStr("9999");
                    EhllapiWrapper.SendStr("@2");


                    EhllapiWrapper.SetCursorPos(244);
                    EhllapiWrapper.ReadScreen(244, 14, out montoAbonado);

                    Thread.Sleep(tiempo);

                    if (montoAbonado != null)
                    {
                        Thread.Sleep(tiempo);


                        EhllapiWrapper.SetCursorPos(575);
                        EhllapiWrapper.SendStr(row1Col1);
                        EhllapiWrapper.SendStr("@A@+");
                        EhllapiWrapper.SendStr("@T");

                        //                    Thread.Sleep(tiempo);
                        //                    //MainFrameAdapter.SetTextOnScreen(7, 25, row1Col1);
                        //                    ////MainFrameAdapter.SendKey(PcomKeys.FieldPlus);
                        //                    ////MainFrameAdapter.SendKey(PcomKeys.Tab);                             
                        //                    //MainFrameAdapter.SetTextOnScreen(7, 32, row1Col2);
                        //                    ////MainFrameAdapter.SendKey(PcomKeys.Tab);
                        //                    //MainFrameAdapter.SetTextOnScreen(7, 35, row1Col3);
                        //                    ////MainFrameAdapter.SendKey(PcomKeys.Tab);
                        //                    //MainFrameAdapter.SetTextOnScreen(7, 39, row1Col4);
                        //                    ////MainFrameAdapter.SendKey(PcomKeys.Tab);
                        //                    //MainFrameAdapter.SetTextOnScreen(7, 43, row1Col5);
                        //                    ////MainFrameAdapter.SendKey(PcomKeys.Tab);
                        //                    ////MainFrameAdapter.SetTextOnScreen(22, 2, row1Col0 + " / APLIC CAP  _" + cont.ToString());
                        //                    //MainFrameAdapter.SetTextOnScreen(22, 2, row1Col0 + " / APLIC CAP");
                        //                    //MainFrameAdapter.SendKey(PcomKeys.PF11);
                        //                    //Thread.Sleep(tiempo);
                        //                    //MainFrameAdapter.SetTextOnScreen(12, 19, ".");
                        //                    //MainFrameAdapter.SendKey(PcomKeys.PF3);


                        //                    //string newColumn = "H";
                        //                    //string newRow = "5";
                        //                    //string worksheet2 = "DATA";

                        //                    ////string sql2 = String.Format("UPDATE [{0}$] SET {1}{2}={3}", worksheet2, newColumn, newRow, "Hola");
                        //                    //string commandString = String.Format("UPDATE [{0}${1}{2}:{1}{2}] SET F1='{3}'", worksheet2, newColumn, newRow, 1);
                        //                    //OleDbCommand objCmdSelect = new OleDbCommand(commandString, connection);
                        //                    //objCmdSelect.ExecuteNonQuery();



                        //                    //COMANDOS PARA EL SIGUIENTE REGISTRO
                        //                    //MainFrameAdapter.SendKey(PcomKeys.PF7);
                        //                    //MainFrameAdapter.SetTextOnScreen(22, 7, "2");
                        //                    //MainFrameAdapter.SendKey(PcomKeys.Enter);
                    }


                }

            }

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

        //private void modoantiguo() {
        //    //string con = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\COMPARTIDO_PUBLICO\RPA\TCS RPA V1.2 - Training\Sample\RPA\Main\To.Rpa.AppCapital\bin\Debug\DATA_PROCESO.xlsx;" + @"Extended Properties='Excel 8.0;HDR=Yes;'";
        //    //string con = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\LOG\DATA_PROCESO.xlsx;Extended Properties='Excel 8.0;HDR=Yes;'";
        //    string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LOG\DATA_PROCESO.xlsx;Extended Properties='Excel 12.0 Xml;HDR=Yes;'";


        //    using (OleDbConnection connection = new OleDbConnection(con))
        //    {
        //        connection.Open();
        //        OleDbCommand command = new OleDbCommand("select * from [DATA$]", connection);
        //        int cont = 0;
        //        using (OleDbDataReader dr = command.ExecuteReader())
        //        {
        //            while (dr.Read())
        //            {
        //                cont = cont + 1;
        //                row1Col0 = dr[0].ToString();
        //                row1Col1 = dr[1].ToString();
        //                row1Col2 = dr[2].ToString();
        //                row1Col3 = dr[3].ToString();
        //                row1Col4 = dr[4].ToString();
        //                row1Col5 = "2908070001210000";//dr[5].ToString();

        //                //MainFrameAdapter.SetTextOnScreen(15, 47, row1Col0);
        //                //MainFrameAdapter.SetTextOnScreen(17, 47, "9999");
        //                //MainFrameAdapter.SendKey(PcomKeys.PF2);
        //                Thread.Sleep(tiempo);
        //                //montoAbonado = MainFrameAdapter.GetScreenText(4, 4, 14);

        //                if (montoAbonado != null)
        //                {
        //                    Thread.Sleep(tiempo);
        //                    //MainFrameAdapter.SetTextOnScreen(7, 25, row1Col1);
        //                    ////MainFrameAdapter.SendKey(PcomKeys.FieldPlus);
        //                    ////MainFrameAdapter.SendKey(PcomKeys.Tab);                             
        //                    //MainFrameAdapter.SetTextOnScreen(7, 32, row1Col2);
        //                    ////MainFrameAdapter.SendKey(PcomKeys.Tab);
        //                    //MainFrameAdapter.SetTextOnScreen(7, 35, row1Col3);
        //                    ////MainFrameAdapter.SendKey(PcomKeys.Tab);
        //                    //MainFrameAdapter.SetTextOnScreen(7, 39, row1Col4);
        //                    ////MainFrameAdapter.SendKey(PcomKeys.Tab);
        //                    //MainFrameAdapter.SetTextOnScreen(7, 43, row1Col5);
        //                    ////MainFrameAdapter.SendKey(PcomKeys.Tab);
        //                    ////MainFrameAdapter.SetTextOnScreen(22, 2, row1Col0 + " / APLIC CAP  _" + cont.ToString());
        //                    //MainFrameAdapter.SetTextOnScreen(22, 2, row1Col0 + " / APLIC CAP");
        //                    //MainFrameAdapter.SendKey(PcomKeys.PF11);
        //                    //Thread.Sleep(tiempo);
        //                    //MainFrameAdapter.SetTextOnScreen(12, 19, ".");
        //                    //MainFrameAdapter.SendKey(PcomKeys.PF3);


        //                    //string newColumn = "H";
        //                    //string newRow = "5";
        //                    //string worksheet2 = "DATA";

        //                    ////string sql2 = String.Format("UPDATE [{0}$] SET {1}{2}={3}", worksheet2, newColumn, newRow, "Hola");
        //                    //string commandString = String.Format("UPDATE [{0}${1}{2}:{1}{2}] SET F1='{3}'", worksheet2, newColumn, newRow, 1);
        //                    //OleDbCommand objCmdSelect = new OleDbCommand(commandString, connection);
        //                    //objCmdSelect.ExecuteNonQuery();



        //                    //COMANDOS PARA EL SIGUIENTE REGISTRO
        //                    //MainFrameAdapter.SendKey(PcomKeys.PF7);
        //                    //MainFrameAdapter.SetTextOnScreen(22, 7, "2");
        //                    //MainFrameAdapter.SendKey(PcomKeys.Enter);
        //                }
        //            }
        //        }
        //    }
        //}

        public void DoUIValidation()
        {
            throw new NotImplementedException();
        }

        public void ExtractElements()
        {
            throw new NotImplementedException();
        }
    }
}
