using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using EHLLAPI;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Automation;

namespace To.AtNinjas.Util
{
    public class Methods
    {
        public static int GetWSPosition(int x, int y)
        {
            return (((x - 1) * 80) + y);
        }

        public static void LogProceso(string sMensaje)
        {
            string dirLog = ConfigurationManager.AppSettings["DirLog"].ToString().Trim();
            string ArcExt = ConfigurationManager.AppSettings["ArcExt"].ToString().Trim();

            DateTime fechaHora = DateTime.Now;
            string fileName = String.Format("{0:ddMMyyyy}", fechaHora);

            string path = dirLog + fileName + ArcExt;
            try
            {
                if (File.Exists(path))
                {
                    File.AppendAllText(path, DateTime.Now.ToString() + " | " + sMensaje.Trim() + " \r\n");
                }
                else
                {
                    using (var file = File.Create(path))
                    {
                        file.Close();
                    }
                    File.AppendAllText(path, DateTime.Now.ToString() + " | " + sMensaje.Trim() + " \r\n");

                }
                Console.WriteLine(sMensaje);
            }
            catch { }
        }
        //TODO: en parametro estaba this, es método de extensión, investigar
        public static Bitmap cropAtRect(Bitmap b, Rectangle r)
        {
            Bitmap nb = new Bitmap(r.Width, r.Height);
            Graphics g = Graphics.FromImage(nb);
            g.DrawImage(b, -r.X, -r.Y);
            return nb;
        }

        public static void Sleep()
        {
            Thread.Sleep(Convert.ToInt32(ConfigurationManager.AppSettings["Sleep"]));
        }

        /*  public static void Sleep_2()
         {
             Thread.Sleep(Convert.ToInt32(ConfigurationManager.AppSettings["Sleep_2"]));
         }*/

        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern bool ShowWindow(IntPtr hwnd, int nCmdShow);


        public static string repairAmount(string amount)
        {
            string newAmount = Regex.Replace(amount, @"[^\d.,]", String.Empty);

            MatchCollection match = Regex.Matches(newAmount, @"(?:\d+[^_]\d)+\w");

            if (match.Count > 0)
            {
                newAmount = match[0].Value;
            }

            decimal result = 0.0m;
            decimal.TryParse(newAmount, out result);

            if (result != 0.0m)
            {
                return newAmount;
            }
            return amount;
        }


        public static Boolean OpenWS(string EmulatorURL, string SessionId, string Username, string Password, out int Id)
        {
            Console.Write("Inicializando IBS... ");
            string IsShowWindow = ConfigurationManager.AppSettings["ShowWindow"].ToString().Trim();
            string Session = ConfigurationManager.AppSettings["SessionId"].ToString().Trim();
            Process[] pcscmProcesses = Process.GetProcessesByName("pcsws");
            if (pcscmProcesses.Length == 0)
            {
                Process.Start("taskkill", "/F /IM [pcsws].exe");
                Process.Start("taskkill", "/F /IM [pcscm].exe");
            }
            //Funciona correcto
            Process px = Process.Start(EmulatorURL);
            //Process px = new Process();
            //try
            //{
            //    px.StartInfo = new ProcessStartInfo(EmulatorURL);
            //    px.StartInfo.CreateNoWindow = true;
            //    px.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
            //    px.Start();

            //} catch (Exception e)
            //{
            //    LogProceso(e.ToString());
            //}

            //Thread.Sleep(4000);
            int _i = 0;
            Id = 0;
            do
            {
                try
                {
                    pcscmProcesses = Process.GetProcessesByName("pcsws");
                    Id = pcscmProcesses.ToList().Find(x => x.MainWindowTitle.Contains(Session)).Id;
                    _i = 1;
                    if (IsShowWindow.Equals("SI"))
                    {
                        ShowWindow(Process.GetProcessById(Id).MainWindowHandle, 2);
                    }
                }
#pragma warning disable CS0168 // La variable 'e' se ha declarado pero nunca se usa
                catch (Exception e)
#pragma warning restore CS0168 // La variable 'e' se ha declarado pero nunca se usa
                {
                    //throw new Exception("No se inició correctamente la sesión B");
                    Thread.Sleep(1000);
                }
            } while (_i == 0);

            Boolean isOpenWS = false;
            string validationPass;
            int i = 0;
            do
            {
                try
                {
                    AutomationElement _0 = AutomationElement.FromHandle(px.MainWindowHandle);
                    AutomationElementCollection _0_Descendants_1 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "ID de usuario:"));
                    ValuePattern etb = _0_Descendants_1[1].GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                    etb.SetValue(Username);

                    AutomationElementCollection _0_Descendants_2 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Contraseña:"));
                    ValuePattern etb2 = _0_Descendants_2[1].GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                    etb2.SetValue(Password);

                    AutomationElementCollection _0_Descendants_3 = _0.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Aceptar"));

                    var invokePattern = _0_Descendants_3[0].GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                    invokePattern.Invoke();
                }
#pragma warning disable CS0168 // La variable 'e' se ha declarado pero nunca se usa
                catch (Exception e)
#pragma warning restore CS0168 // La variable 'e' se ha declarado pero nunca se usa
                {
                }


                EhllapiWrapper.Connect(SessionId);
                EhllapiWrapper.Wait();
                EhllapiWrapper.SetCursorPos(GetWSPosition(6, 17));
                EhllapiWrapper.Wait();
                EhllapiWrapper.ReadScreen(GetWSPosition(6, 17), 7, out validationPass);
                EhllapiWrapper.Wait();
                try
                {
                    if (validationPass.Substring(0, 7) == "Usuario")
                    {
                        var ValidationMsg = "";

                        EhllapiWrapper.Wait();
                        EhllapiWrapper.Connect(SessionId);
                        EhllapiWrapper.Wait();
                        EhllapiWrapper.SetCursorPos(GetWSPosition(6, 53));
                        EhllapiWrapper.Wait();
                        EhllapiWrapper.SendStr(Username);
                        EhllapiWrapper.Wait();
                        EhllapiWrapper.SetCursorPos(GetWSPosition(7, 53));
                        EhllapiWrapper.Wait();
                        EhllapiWrapper.SendStr(Password);
                        EhllapiWrapper.Wait();
                        EhllapiWrapper.SendStr("@E");
                        EhllapiWrapper.Wait();
                        ValidationMsg = "";
                        EhllapiWrapper.SetCursorPos(GetWSPosition(1, 26));
                        EhllapiWrapper.Wait();
                        EhllapiWrapper.ReadScreen(GetWSPosition(1, 26), 31, out ValidationMsg);
                        EhllapiWrapper.Wait();
                        if (ValidationMsg.Substring(0, 31).Equals("Visualizar Mensajes de Programa"))
                        {
                            EhllapiWrapper.SendStr("@E");
                            EhllapiWrapper.Wait();
                            isOpenWS = true;
                        }
                        else
                        {
                            Console.Write("OK\n");
                            isOpenWS = true;
                        }
                        i = 1;
                    }
                }
#pragma warning disable CS0168 // La variable 'e' se ha declarado pero nunca se usa
                catch (Exception e)
#pragma warning restore CS0168 // La variable 'e' se ha declarado pero nunca se usa
                {
                }
            } while (i == 0);
            return isOpenWS;
        }

        public static void GetMenuProgram(String Menu)
        {
            var options = Menu.Split(',');
            LogProceso("Invocando el programa del menú: " + Menu);
            foreach (var option in options)
            {
                EhllapiWrapper.SetCursorPos(GetWSPosition(22, 7));
                EhllapiWrapper.Wait();
                EhllapiWrapper.SendStr(option);
                EhllapiWrapper.Wait();
                EhllapiWrapper.SendStr("@E");
                EhllapiWrapper.Wait();
            }
        }

        //public static void UpdateCell(SpreadsheetDocument document, string text,
        //    uint rowIndex, string columnName)
        //{
        //    // Open the document for editing.
        //    WorksheetPart worksheetPart =
        //          GetWorksheetPartByName(document, "Hoja2");

        //    if (worksheetPart != null)
        //    {
        //        Cell cell = GetCell(worksheetPart.Worksheet,
        //                                 columnName, rowIndex);

        //        cell.CellValue = new CellValue(text);

        //        //TODO: VER SI SE COMENTA
        //        cell.DataType =
        //            new EnumValue<CellValues>(CellValues.Number);

        //        // Save the worksheet.
        //        worksheetPart.Worksheet.Save();
        //    }
        //}

        public static void UpdateCell(string fileName, string text,
            uint rowIndex, string columnName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
            {
                // Open the document for editing.
                WorksheetPart worksheetPart =
                      //TODO:parametrizar hoja
                      GetWorksheetPartByName(document, "Hoja1");

                if (worksheetPart != null)
                {
                    Cell cell = GetCell(worksheetPart.Worksheet,
                                             columnName, rowIndex);

                    cell.CellValue = new CellValue(text);

                    //TODO: VER SI SE COMENTA
                    cell.DataType =
                        new EnumValue<CellValues>(CellValues.Number);

                    // Save the worksheet.
                    worksheetPart.Worksheet.Save();
                }
            }
        }

        private static WorksheetPart
             GetWorksheetPartByName(SpreadsheetDocument document,
             string sheetName)
        {
            IEnumerable<Sheet> sheets =
               document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
               Elements<Sheet>().Where(s => s.Name == sheetName);

            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.

                return null;
            }

            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)
                 document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;

        }

        // Given a worksheet, a column name, and a row index, 
        // gets the cell at the specified column and 
        private static Cell GetCell(Worksheet worksheet,
                  string columnName, uint rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);

            if (row == null)
                return null;

            return row.Elements<Cell>().Where(c => string.Compare
                   (c.CellReference.Value, columnName +
                   rowIndex, true) == 0).First();
        }

        // Given a worksheet and a row index, return the row.
        private static Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().
              Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }

        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

    }
}
