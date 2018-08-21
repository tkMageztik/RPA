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
using System.Configuration;
using To.AtNinjas.Util;
using System.Reflection;

namespace To.Rpa.AppCapital.UserInterfaces
{
    public class ScreenLogin
    {
        //private Stopwatch objWatch = new Stopwatch();

        public void CheckCompleteLoad()
        {

        }

        private string User { get; set; }
        private string Password { get; set; }
        private string Template { get; set; }

        private void LoadAccess()
        {
            User = ConfigurationManager.AppSettings["MFUser"];
            Password = ConfigurationManager.AppSettings["MFPass"];
        }

        private void LoadConfig()
        {
            Template = ConfigurationManager.AppSettings["MFTemplate"];
        }

        public void DoActivities()
        {
            //EntityGlobal.lastScreenName = "Login";

            CheckCompleteLoad();
            // ExtractElements();
            // DoUIValidation();
            DoOperations();

        }

        public void DoOperations()
        {
            //TODO: Depende de que la sesión "A" esté libre

            var iConnection = EhllapiWrapper.Connect("A").ToString();

            Debug.WriteLine("Valor de conexión: " + iConnection);

            string lectura = null;
            if (iConnection == "0")
            {
                Methods.Sleep();
                EhllapiWrapper.Wait();
                LoadAccess();
                LoadConfig();
                EhllapiWrapper.SetCursorPos(453);
                EhllapiWrapper.SendStr(User);

                EhllapiWrapper.SetCursorPos(533);
                EhllapiWrapper.SendStr(Password);

                EhllapiWrapper.SendStr("@E");
                EhllapiWrapper.Wait();

                EhllapiWrapper.SetCursorPos(162);

                EhllapiWrapper.ReadScreen(162, 239, out lectura);

                //if (lectura.IndexOf("La cola de mensajes BFPJUARUI está asignada a otro trabajo") != 0)
                if (lectura.IndexOf("está asignada a otro trabajo") != -1)
                {
                    EhllapiWrapper.SendStr("@E");
                    EhllapiWrapper.Wait();
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
                EhllapiWrapper.Wait();

                FileInfo existingFile = new FileInfo(Template);

                using (ExcelPackage package = new ExcelPackage(existingFile))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];

                    int row = 2;
                    //TODO: aqui mover capital... .
                    string prestamo = "", banco, agencia, moneda, ctaTramite, cta, estado, capital, error, deduccion;
                    decimal monto, tryParse;

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

                            Console.WriteLine("prestamo " + prestamo);
                            Console.WriteLine("monto " + monto);
                            Console.WriteLine("ctaTramite " + ctaTramite);
                            Console.WriteLine("cta " + cta);

                            EhllapiWrapper.SetCursorPos(1167);
                            EhllapiWrapper.SendStr("@F");
                            EhllapiWrapper.SendStr(prestamo);
                            Console.ReadKey();
                            EhllapiWrapper.SetCursorPos(1327);
                            EhllapiWrapper.SendStr("9999");
                            EhllapiWrapper.SendStr("@2");
                            EhllapiWrapper.Wait();

                            capital = null;

                            EhllapiWrapper.SetCursorPos(244);
                            EhllapiWrapper.ReadScreen(244, 14, out capital);
                            capital = capital.Substring(0, 14);

                            Console.WriteLine("capiii " + row + " " + capital);
                            Console.ReadKey();

                            //TODO: para que sirve esta validación? (si el contrato/préstamo no existe
                            if (Decimal.TryParse(capital, out tryParse))
                            {
                                Console.ReadKey();
                                Console.WriteLine("entro JRDC " + row);
                                //blanquear
                                //(Pag/Cancela) / monto
                                EhllapiWrapper.SetCursorPos(575);
                                EhllapiWrapper.SendStr("@F");

                                //Bn / Banco
                                EhllapiWrapper.SetCursorPos(592);
                                EhllapiWrapper.SendStr("@F");

                                //Suc / Sucursal
                                EhllapiWrapper.SetCursorPos(595);
                                EhllapiWrapper.SendStr("@F");

                                //Mda / Moneda 
                                EhllapiWrapper.SetCursorPos(599);
                                EhllapiWrapper.SendStr("@F");

                                //Cuenta Contable / Cuenta Trámite
                                EhllapiWrapper.SetCursorPos(603);
                                EhllapiWrapper.SendStr("@F");

                                //Cuenta
                                EhllapiWrapper.SetCursorPos(620);
                                EhllapiWrapper.SendStr("@F");

                                EhllapiWrapper.SetCursorPos(575);
                                EhllapiWrapper.SendStr(monto.ToString());
                                EhllapiWrapper.SendStr("@A@+");
                                EhllapiWrapper.Wait();
                                //EhllapiWrapper.SendStr("@T");

                                if (!ctaTramite.Equals(""))
                                {
                                    EhllapiWrapper.SetCursorPos(603);
                                    EhllapiWrapper.SendStr(ctaTramite);

                                    banco = worksheet.Cells[row, 3].Value.ToString();
                                    agencia = worksheet.Cells[row, 4].Value.ToString();
                                    moneda = worksheet.Cells[row, 5].Value.ToString();

                                    EhllapiWrapper.SetCursorPos(592);
                                    EhllapiWrapper.SendStr(banco);

                                    EhllapiWrapper.SetCursorPos(595);
                                    EhllapiWrapper.SendStr(agencia);

                                    EhllapiWrapper.SetCursorPos(599);
                                    EhllapiWrapper.SendStr(moneda);
                                }
                                else
                                {
                                    //TODO: CUENTA
                                    EhllapiWrapper.SetCursorPos(620);
                                    EhllapiWrapper.SendStr(cta);
                                }

                                EhllapiWrapper.SetCursorPos(1682);
                                EhllapiWrapper.SendStr("@F");
                                EhllapiWrapper.SendStr(prestamo + " / APLIC CAP");
                                EhllapiWrapper.SendStr("@E");
                                EhllapiWrapper.Wait();

                                /********************************************************/
                                EhllapiWrapper.ReadScreen(244, 14, out error);
                                if (error == "mensaje de error")
                                {
                                    EhllapiWrapper.SendStr("@E");
                                    EhllapiWrapper.Wait();

                                    EhllapiWrapper.ReadScreen(1467, 4, out deduccion);

                                    if (!deduccion.Trim().Equals(""))
                                    {
                                        //TODO: UBICAR BIEN DEDUCCIONES
                                        EhllapiWrapper.SetCursorPos(1455);
                                        EhllapiWrapper.SendStr("@c");
                                        EhllapiWrapper.Wait();
                                    }
                                }
                                /********************************************************/

                                //EhllapiWrapper.SetCursorPos(328);
                                EhllapiWrapper.ReadScreen(328, 57, out error);

                                if (error.Equals("Error en Número de Cuenta Contable o Falta, Favor Revisar"))
                                {
                                    Methods.LogProceso("ENTRÓ!");
                                    EhllapiWrapper.SendStr("@E");
                                    EhllapiWrapper.Wait();

                                    EhllapiWrapper.ReadScreen(1455, 16, out deduccion);
                                    Methods.LogProceso("Deducción obtenida " + deduccion);

                                    //if (deduccion.Trim() == "existen deducciones")
                                    if (!deduccion.Trim().Equals(""))
                                    {
                                        //TODO: UBICAR BIEN DEDUCCIONES
                                        EhllapiWrapper.SetCursorPos(1455);
                                        EhllapiWrapper.SendStr("@c");
                                        EhllapiWrapper.Wait();
                                    }
                                } 

                                //TODO: CAPITAL
                                //MetodoRecursivo(monto);


                                EhllapiWrapper.SendStr("@b");
                                EhllapiWrapper.Wait();

                                EhllapiWrapper.SetCursorPos(899);
                                EhllapiWrapper.SendStr(".");
                                EhllapiWrapper.SendStr("@3");
                                EhllapiWrapper.Wait();

#if DEBUG
                                EhllapiWrapper.SendStr("@3");
                                EhllapiWrapper.Wait();
#endif

                                EhllapiWrapper.SetCursorPos(1167);
                                EhllapiWrapper.SendStr("@F");
                                EhllapiWrapper.Wait();
                            }
                            else
                            {
                                //MARCAR ERROR 
                                EhllapiWrapper.ReadScreen(1526, 70, out error);
                                worksheet.SetValue(row, 8, "ERROR - " + error);
                            }

                            if (worksheet.Cells[row + 1, 1].Value == null)
                            {
                                prestamo = "";
                            }
                        }
                        catch (Exception exc)
                        {
                            //No está funcionando escribir en el excel.
                            worksheet.SetValue(row, 8, "ERROR");
                            Methods.LogProceso(" ERROR: " + exc.StackTrace + " " + MethodBase.GetCurrentMethod().Name);
                            EhllapiWrapper.SendStr("@E");
                            EhllapiWrapper.Wait();
                        }
                        finally { row++; }

                    } while (!prestamo.Equals(""));

                }
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
