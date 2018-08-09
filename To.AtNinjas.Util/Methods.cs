using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace To.AtNinjas.Util
{
    public class Methods
    {
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
