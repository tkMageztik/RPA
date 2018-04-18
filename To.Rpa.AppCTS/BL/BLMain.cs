using NHunspell;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Tesseract;
using To.Rpa.Util;
using To.Rpa.AppCTS.BE;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using S = DocumentFormat.OpenXml.Spreadsheet.Sheets;
using E = DocumentFormat.OpenXml.OpenXmlElement;
using A = DocumentFormat.OpenXml.OpenXmlAttribute;
using DocumentFormat.OpenXml;

namespace To.Rpa.AppCTS.BL
{
    public class BLMain
    {
        private FileInfo[] OriginFiles { get; set; }
        private string OriginSharedDirectory { get; set; }
        private string OriginDirectory { get; set; }
        private string ScrapsDirectory { get; set; }
        private string TemplatesDirectory { get; set; }
        private List<BECTSFormat> ApprovedFormats { get; set; }
        private List<BECTSFormat> DisapprovedFormats { get; set; }
        private List<BECTSImagePages> col2ImagesPagesTemp { get; set; }
        //private string col2FileNameTemp { get; set; }
        private bool foundHeader { get; set; }
        private string tesseractLanguage { get; set; }

        public BLMain()
        {
            OriginSharedDirectory = ConfigurationManager.AppSettings["OriginSharedDirectory"];
            OriginDirectory = ConfigurationManager.AppSettings["OriginDirectory"];
            ScrapsDirectory = ConfigurationManager.AppSettings["ScrapsDirectory"];
            TemplatesDirectory = ConfigurationManager.AppSettings["TemplatesDirectory"];
            tesseractLanguage = ConfigurationManager.AppSettings["TesseractLanguage"];

            //TODO, no debería ser lazy load... .
            ApprovedFormats = new List<BECTSFormat>();
            DisapprovedFormats = new List<BECTSFormat>();
            col2ImagesPagesTemp = new List<BECTSImagePages>();
        }

        public void DoActivities()
        {
            //GetPdfFromSharedDirectory();
            //GetPdfFromMailBox(); MUEVE A \WORKSPACE\ORIGENES
            //CreateFoldersPerCTSFormat();
            //ExtractImagesFromPdf();
            //ImproveImagesFromPdf();
            GetApprovedFormats();

            GetScrapsDataImages();
            //GetScrapsDataImagesFull();
            //GetHeaderDataFromScraps();
            GetDataFromScraps();
            //MoveOriginsNoProcess(); lvl 01
        }
        //TODO: por que los thumbnails de windows muestran imagenes que no corresponden a la realidad??
        private int getFormatCode(string fileName)
        {
            //el formato es un guión (-), seguido de un número, se podría hacer más inteligente.
            string[] scraps = fileName.Split('-');

            int formatCode;
            int.TryParse(Path.GetFileNameWithoutExtension(scraps[scraps.Length - 1].Trim()), out formatCode);
            return formatCode;
        }

        private void CreateFoldersPerCTSFormat()
        {
            //TODO se puede cambiar por la lista de la variable de clasee...
            DirectoryInfo diOrigin = new DirectoryInfo(OriginDirectory);
            DirectoryInfo diScraps = new DirectoryInfo(ScrapsDirectory);
            DirectoryInfo diScrapsChild = null;

            string fileName = "";
            string directoryFullName = "";
            string directoryName = "";

            List<string> lstFilePaths = new List<string>();

            //FileInfo[] files = OriginFiles;
            FileInfo[] files = diOrigin.GetFiles();

            int totalFiles = files.Length;

            for (int i = 0; i < totalFiles; i++)
            {
                //TODO: Separado en config
                //string[] scraps = files[i].Name.Split('-');

                //int formatCode;
                //int.TryParse(Path.GetFileNameWithoutExtension(scraps[scraps.Length - 1].Trim()), out formatCode);

                int formatCode = getFormatCode(files[i].Name);

                if (formatCode != 0)
                {
                    //TODO: Checar si es que hay carpeta repetido cuidado se sobreescribirá..
                    diScrapsChild = new DirectoryInfo(ScrapsDirectory + formatCode + @"\");
                    diScrapsChild.Create();
                    diScrapsChild.Refresh();

                    diScrapsChild = new DirectoryInfo(ScrapsDirectory + formatCode + @"\ORIGINAL\");
                    diScrapsChild.Create();
                    diScrapsChild.Refresh();

                    diScrapsChild = new DirectoryInfo(ScrapsDirectory + formatCode + @"\PLANTILLA\");
                    diScrapsChild.Create();
                    diScrapsChild.Refresh();

                    File.Move(OriginDirectory + files[i].Name,
                        ScrapsDirectory + formatCode + @"\ORIGINAL\" + files[i].Name);
                }
            }
        }


        private void ImproveImagesFromPdf()
        {
            throw new NotImplementedException();
        }

        private void ExtractImagesFromPdf()
        {
            throw new NotImplementedException();
        }

        private void GetPdfFromMailBox()
        {

        }
        //TODO guardar nombre de los archivos en la primera lectura sin criterios ... para luego moverlos... .,los que
        //quedan es porque no se procesaron.
        private void GetPdfFromSharedDirectory()
        {
            DirectoryInfo diOrigin = new DirectoryInfo(OriginSharedDirectory);

            OriginFiles = diOrigin.GetFiles();

            int totalFiles = OriginFiles.Length;

            for (int i = 0; i < totalFiles; i++)
            {
                string formatCode = getFormatCode(OriginFiles[i].Name) + "";

                if (!formatCode.Equals(0))
                {
                    DirectoryInfo diDestiny = new DirectoryInfo(OriginDirectory);

                    FileInfo[] destinyFiles = diDestiny.GetFiles("*-*" + formatCode + "*");

                    int length = destinyFiles.Length;

                    if (length > 0)
                    {
                        File.Move(diOrigin.FullName + OriginFiles[i].Name,
                            diDestiny.FullName + Path.GetFileNameWithoutExtension(OriginFiles[i].Name) + "_" + length + Path.GetExtension(OriginFiles[i].Name));
                    }
                    else
                    {
                        File.Move(diOrigin.FullName + OriginFiles[i].Name,
                            diDestiny.FullName + OriginFiles[i].Name);
                    }
                }
            }
        }

        private void GetApprovedFormats()
        {
            DirectoryInfo di = new DirectoryInfo(ScrapsDirectory);
            string fileName = "";
            string directoryFullName = "";
            string directoryName = "";
            bool approved = false;
            List<string> lstPathDirectories = new List<string>();

            DirectoryInfo[] directories = di.GetDirectories();
            //FileInfo[] files = di.GetFiles();

            int totalDirectories = directories.Length;


            for (int i = 0; i < totalDirectories; i++)
            {
                List<BECTSImagePages> lstBECTSImagePages = new List<BECTSImagePages>();
                foundHeader = false;

                directoryFullName = directories[i].FullName;
                directoryName = directories[i].Name;
                //lstPathDirectories.Add(directoryName);

                //di.GetFiles()
                di = new DirectoryInfo(directoryFullName + @"\ORIGINAL\");
                //di = new DirectoryInfo(directoryFullName);
                FileInfo[] files = di.GetFiles("base_" + directoryName + "*");

                int totalArchivos = files.Length;

                for (int o = 0; o < totalArchivos; o++)
                {

                    //COPIA LA IMAGEN DEJA EL ORIGINAL, evalua los intentos por si algún motivo el archivo esté siendo usando.
                    int attemps = 0;
                    while (attemps < Int32.Parse(ConfigurationManager.AppSettings["Attempts"]))
                    {
                        try
                        {
                            //int attemps = 0;
                            File.Copy(files[o].FullName, files[o].FullName.Replace(@"\ORIGINAL", ""));
                            attemps = Int32.Parse(ConfigurationManager.AppSettings["Attempts"]);
                        }
                        catch (Exception exc) { attemps++; }
                    }

                    for (int u = 0; u < 3; u++)
                    {
                        using (Bitmap image = new Bitmap(files[o].FullName.Replace(@"\ORIGINAL", "")))
                        {
                            Bitmap newImage = null;

                            //ANTES EVALUABA POR NOMBRES Y APELLIDOS
                            //COLUMNA 2 - NOMBRES Y APELLIDOS
                            //newImage = Methods.cropAtRect(image, new Rectangle(312, 1494, 400, image.Height - 1494));
                            //newImage = Methods.cropAtRect(image, new Rectangle(450, 1494, 850, image.Height - 1494));
                            //fileName = files[o].DirectoryName.Replace(@"\ORIGINAL", "") + @"\" + Path.GetFileNameWithoutExtension(files[o].Name) + "_1" + Path.GetExtension(files[0].Name);
                            //newImage.Save(fileName);

                            //AHORA SE EVALÚA POR NRO DOCUMENTO
                            //COLUMNA 4 - NRO DOCUMENTO
                            if (!foundHeader)
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(1570, 1494, 350, image.Height - 1494));
                            }
                            else
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(1570, 10, 350, image.Height - 10));
                            }
                            //fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_3" + Path.GetExtension(imagePages.CodePage);
                            fileName = files[o].DirectoryName.Replace(@"\ORIGINAL", "") + @"\" + Path.GetFileNameWithoutExtension(files[o].Name) + "_4" + Path.GetExtension(files[0].Name);
                            newImage.Save(fileName);

                            //if (IsDataImage(fileName) || (files[o].Name == "base_4023_1.png" && u == 1))
                            //{
#if DEBUG
                            if (IsDataImage(fileName))
                            {
#else
                            if (IsDataImage(fileName))
                            {
#endif
                                //SUPONGAMOS QUE SI ES EL SEGUNDO, YA TIENE QUE VENIR DESDE ARRIBITA.
                                approved = true;
                                //GUARDO EL FILENAME PATH DEL SCRAP DE LA COLUMNA NRO DOCUMENTO EN UNA TEMP PARA LUEGO GUARDALO EN ORDEN
                                //col2FileNameTemp = fileName;

                                if (!foundHeader)
                                {
                                    lstBECTSImagePages.Add(new BECTSImagePages { CodePage = files[o].Name, isHeader = true });
                                    col2ImagesPagesTemp.Add(new BECTSImagePages { CodePage = files[o].Name, isHeader = true, ImagePagesScraps = new List<string>() { fileName } });
                                    foundHeader = true;
                                }
                                else
                                {
                                    lstBECTSImagePages.Add(new BECTSImagePages { CodePage = files[o].Name });
                                    col2ImagesPagesTemp.Add(new BECTSImagePages { CodePage = files[o].Name, ImagePagesScraps = new List<string>() { fileName } });
                                }
                                break;
                            }
                            else
                            {
                                ImageRotate(image, files[o].FullName.Replace(@"\ORIGINAL", ""));
                            }

                        }
                    }

                }

                if (approved)
                {
                    ApprovedFormats.Add(new BECTSFormat
                    {
                        FormatCode = int.Parse(directoryName),
                        LstBECTSImagePages = lstBECTSImagePages
                    });
                }
                else
                {
                    DisapprovedFormats.Add(new BECTSFormat
                    {
                        FormatCode = int.Parse(directoryName),
                        LstBECTSImagePages = lstBECTSImagePages
                    });
                }

            }
        }

        private void GetScrapsDataImages()
        {
            foreach (BECTSFormat formats in ApprovedFormats)
            {
                foreach (BECTSImagePages imagePages in formats.LstBECTSImagePages)
                {
                    string fileName = ScrapsDirectory + formats.FormatCode + @"\" + imagePages.CodePage;
                    using (Bitmap image = new Bitmap(fileName))
                    {

                        Bitmap newImage = null;
                        //COLUMNA 2 - NOMBRES Y APELLIDOS
                        if (imagePages.isHeader)
                        {
                            newImage = Methods.cropAtRect(image, new Rectangle(450, 1494, 850, image.Height - 1494));
                        }
                        else
                        {
                            newImage = Methods.cropAtRect(image, new Rectangle(450, 10, 850, image.Height - 10));
                        }
                        fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_2" + Path.GetExtension(imagePages.CodePage);
                        newImage.Save(fileName);
                        imagePages.ImagePagesScraps.Add(fileName);

                        //COLUMNA 3 - TIPO DOCUMENTO
                        if (imagePages.isHeader)
                        {
                            newImage = Methods.cropAtRect(image, new Rectangle(1320, 1494, 250, image.Height - 1494));
                        }
                        else
                        {
                            newImage = Methods.cropAtRect(image, new Rectangle(1320, 10, 250, image.Height - 10));
                        }
                        fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_3" + Path.GetExtension(imagePages.CodePage);
                        newImage.Save(fileName);
                        imagePages.ImagePagesScraps.Add(fileName);

                        //COLUMNA 4 - NRO DOCUMENTO
                        //imagePages.ImagePagesScraps.Add(col2FileNameTemp);
                        imagePages.ImagePagesScraps.Add(col2ImagesPagesTemp.Where(x => x.CodePage == imagePages.CodePage).Select(x => x.ImagePagesScraps[0]).First());

                        //newImage = Methods.cropAtRect(image, new Rectangle(1570, 1494, 350, image.Height - 1494));
                        //fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_3" + Path.GetExtension(imagePages.CodePage);
                        //newImage.Save(fileName);
                        //imagePages.ImagePagesScraps.Add(fileName);

                        //COLUMNA 5 - NRO CUENTA
                        if (imagePages.isHeader)
                        {
                            newImage = Methods.cropAtRect(image, new Rectangle(1920, 1494, 400, image.Height - 1494));
                        }
                        else
                        {
                            newImage = Methods.cropAtRect(image, new Rectangle(1920, 10, 400, image.Height - 10));
                        }
                        fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_5" + Path.GetExtension(imagePages.CodePage);
                        newImage.Save(fileName);
                        imagePages.ImagePagesScraps.Add(fileName);

                        //COLUMNA 6 - CUENTA > MONEDA
                        if (imagePages.isHeader)
                        {
                            newImage = Methods.cropAtRect(image, new Rectangle(2420, 1494, 220, image.Height - 1494));
                        }
                        else
                        {
                            newImage = Methods.cropAtRect(image, new Rectangle(2420, 10, 220, image.Height - 10));
                        }
                        fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_6" + Path.GetExtension(imagePages.CodePage);
                        newImage.Save(fileName);
                        imagePages.ImagePagesScraps.Add(fileName);

                        //COLUMNA 7 - ABONO > MONTO
                        if (imagePages.isHeader)
                        {
                            newImage = Methods.cropAtRect(image, new Rectangle(2660, 1494, 380, image.Height - 1494));
                        }
                        else
                        {
                            newImage = Methods.cropAtRect(image, new Rectangle(2660, 10, 380, image.Height - 10));
                        }
                        fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_7" + Path.GetExtension(imagePages.CodePage);
                        newImage.Save(fileName);
                        imagePages.ImagePagesScraps.Add(fileName);

                        //COLUMNA 8 - ABONO > MONEDA
                        if (imagePages.isHeader)
                        {
                            newImage = Methods.cropAtRect(image, new Rectangle(3040, 1494, 220, image.Height - 1494));
                        }
                        else
                        {
                            newImage = Methods.cropAtRect(image, new Rectangle(3040, 10, 220, image.Height - 10));
                        }
                        fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_8" + Path.GetExtension(imagePages.CodePage);
                        newImage.Save(fileName);
                        imagePages.ImagePagesScraps.Add(fileName);

                        //COLUMNA 9 - 4 REMUNERACIONES > MONTO
                        if (imagePages.isHeader)
                        {
                            newImage = Methods.cropAtRect(image, new Rectangle(3360, 1494, 480, image.Height - 1494));
                        }
                        else
                        {
                            newImage = Methods.cropAtRect(image, new Rectangle(3360, 10, 480, image.Height - 10));
                        }
                        fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_9" + Path.GetExtension(imagePages.CodePage);
                        newImage.Save(fileName);
                        imagePages.ImagePagesScraps.Add(fileName);

                        //COLUMNA 10 - 4 REMUNERACIONES > MONEDA
                        if (imagePages.isHeader)
                        {
                            newImage = Methods.cropAtRect(image, new Rectangle(3840, 1494, 220, image.Height - 1494));
                        }
                        else
                        {
                            newImage = Methods.cropAtRect(image, new Rectangle(3840, 10, 220, image.Height - 10));
                        }
                        fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_10" + Path.GetExtension(imagePages.CodePage);
                        newImage.Save(fileName);
                        imagePages.ImagePagesScraps.Add(fileName);

                        //imagePages.ImagePagesScraps = imagePages.ImagePagesScraps.OrderBy(x => x).ToList();
                    }
                }
            }
        }

        private void ImageRotate(Image image, string fileName)
        {
            image.RotateFlip(RotateFlipType.Rotate90FlipXY);
            image.Save(fileName);
        }
        //TODO: falta extraer los que tienen la tabla del tamaño de toda la hoja
        private void GetDataFromScraps()
        {
            foreach (BECTSFormat formats in ApprovedFormats)
            {
                //COPIO PLANTILLA CON NOMBRE DE FORMATO.
                string filename = ScrapsDirectory + formats.FormatCode + @"\plantilla\" + "Plantilla abono CTS" + " - " + formats.FormatCode + ".xlsx";

                //TODO: validar si el archivo existe y marcarlo en el log.
                File.Copy(TemplatesDirectory + "Plantilla abono CTS.xlsx", filename, true);
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(filename, true))
                {
                    uint actualRow = 0;
                    uint lastRowCountImagePage = 0;
                    uint lastRowCountImagePageTemp = 0;

                    foreach (BECTSImagePages imagePages in formats.LstBECTSImagePages)
                    {
                        int initialCol = 2;
                        string col = "";
                        //TODO: el salto de lineas en blanco debe ser parámetro, o hago método que evalue q tanto se parece la cadena al nombre que devuelve la búsqueda por dni en la reniec. o ignoro la columna nombres y apellidos desde el pdf y la pinto desde reniec.
                        //le agregamos más 2  espacios en blanco en caso se salte uno o máximo dos por error.
                        lastRowCountImagePage += lastRowCountImagePageTemp;
                        foreach (string scrap in imagePages.ImagePagesScraps)
                        {
                            //TODO: al terminar de grabar en el excel se dan errores que se pueden autoreparar en el excel. STACKOVERFLOW
                            actualRow = 0;
                            switch (initialCol)
                            {
                                case 2: col = "C"; break;
                                case 3: col = "D"; break;
                                case 4: col = "E"; break;
                                case 5: col = "F"; break;
                                case 6: col = "G"; break;
                                case 7: col = "H"; break;
                                case 8: col = "I"; break;
                                case 9: col = "J"; break;
                                default: col = "K"; break; //10
                            }
                            //TODO:EL TIPO DE DOCUMENTO NO ESTÁ LEYENDO NADA.
                            List<string> dataCol2 = GetData(scrap);
                            foreach (string data in dataCol2)
                            {
                                //TODO: AL PARECER EL PRIMER REGISTRO ES VACIO SIEMPRE, PERO PODRÍA DARSE EL CASO Q EN EL MEDIO TMB SE ENCUENTRE VACIO, SOLO IMPORTA Q NO SALGA VACIO AL MEDIO EN DNI.
                                //row1Col0 = GetCellValue(document, "DATA", "A2");
                                if (data.Trim() != "")
                                {
                                    string dataRepaired = data;

                                    if (col == "E")
                                    {
                                        dataRepaired = repairID("", data);
                                    }

                                    if (col == "F")
                                    {
                                        dataRepaired = repairAccountNumber("", data);
                                    }

                                    if (col == "H" || col == "J")
                                    {
                                        dataRepaired = repairAmount(data);
                                    }

                                    UpdateCell(document, dataRepaired, 25 + actualRow + lastRowCountImagePage, col);
                                    actualRow++;
                                }
                            }
                            initialCol++;

                            if (col == "E")
                            {
                                lastRowCountImagePageTemp = actualRow + 2;
                            }
                        }
                    }
                }
            }
        }

        private string repairID(string names, string ID)
        {
            string newID = Regex.Replace(ID, @"[^\d]", String.Empty).PadLeft(8, '0');

            if (Regex.IsMatch(newID, @"\d{8}"))
            {
                //TODO, si machea regex igual evaluar si existe dni con reniec... 
                return newID;
            }

            //newID = Regex.Match(ID.Trim().Replace(" ", "").PadLeft(8, '0'), @"\d{8}").Value;


            //TODO: speudocódigo con reniec si no matechea con regex.
            //if (newID == "") {

            //    newID = ID.PadLeft(8, '0');

            //    namesReniec = metodo(ID);

            //    if (names similar namesReniec){ return newID }

            //    newID = ID.PadRight(8, '0');
            //    namesReniec = metodo(ID);

            //    if (names similar namesReniec){ return newID }
            //}




            //TODO: podemos pintarlo de rojo si no machea. o escribirlo en el log.
            return ID;
        }

        private string repairAccountNumber(string ID, string account)
        {
            string newAccount = Regex.Replace(account, @"[^\d]", String.Empty).PadLeft(12, '0');
            //Regex.IsMatch(newAccount, @"\d{12}");

            if (Regex.IsMatch(newAccount, @"\d{12}"))
            {
                return newAccount;

                //TODO: speudocódigo con ibs
                //if (newAccount == "") {

                // 1oMuchasCtas= buscarCuentasPorDNI()
                // if (alguna 1oMuchasCtas es similar a account en un 90% o si simplemente existe de manera idéntica) { return newAccount}

            }

            return newAccount;
        }

        private string repairAmount(string amount)
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

        public static void UpdateCell(SpreadsheetDocument document, string text,
            uint rowIndex, string columnName)
        {
            // Open the document for editing.
            WorksheetPart worksheetPart =
                  GetWorksheetPartByName(document, "Hoja2");

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

        //
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

        //TODO: para pasar a algún método 
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

        private List<string> GetData(string scrapPath)
        {
            try
            {
                //TODO: si va a ser solo números ubicar el engine en inglés, es más rápido al parecer... .
                using (var engine = new TesseractEngine(@"./Dictionaries/Teserract", "eng", EngineMode.Default))
                {

                    using (var img = Pix.LoadFromFile(scrapPath))
                    {
                        using (var page = engine.Process(img))
                        {
                            //TODO: parece que el primer registro lo lee como salto de línea.
                            var text = page.GetText();

                            List<string> prev = text.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries).ToList<string>();
                            return prev;
                        }
                    }
                }
            }
            catch { }

            return null;
        }

        private bool IsDataImage(string scrapIDPath)
        {
            if (ValidateIDOnPremise(scrapIDPath) && ValidateIDOnReniec(scrapIDPath))
            {
                return true;
            }
            return false;
        }

        private bool ValidateIDOnPremise(string scrapIDPath)
        {
            List<string> lstIDs = GetData(scrapIDPath);

            //TODO: meter el minimunAccepted en appconfig
            Int16 minimunAccepted = (Int16)(lstIDs.Count / 2);
            Int16 accepted = 0;
            foreach (string ID in lstIDs)
            {
                if (Regex.IsMatch(ID, @"\d{8}"))
                //if (Regex.Match(ID, @"\d{8}").Success)
                {
                    accepted++;
                }
            }

            if (accepted > minimunAccepted)
            {
                //TODO: guardar lo le´´ido para optimziar.
                return true;
            }

            else return false;
        }
        private bool ValidateIDOnReniec(string scrapIDPath)
        {
            return true;
        }

        private bool ValidateIDOnPremise_2()
        {
            return true;
        }

        private void ValidateExample(string filePath)
        {
            var testImagePath = filePath;
            //if (args.Length > 0)
            //{
            //    testImagePath = args[0];
            //}

            try
            {
                using (var engine = new TesseractEngine(@"./Dictionaries/Teserract", "spa", EngineMode.Default))
                {

                    using (var img = Pix.LoadFromFile(testImagePath))
                    {

                        using (var page = engine.Process(img))
                        {
                            //TODO: parece que el primer registro lo lee como salto de línea.
                            var text = page.GetText();

                            List<string> result = text.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries).ToList<string>();

                            for (int i = result.Count - 1; i >= 0; i--)
                            {
                                using (Hunspell hunspell = new Hunspell("Dictionaries/Hunspell/es_ES.aff", "Hunspell_dictionaries/es_ES.dic"))
                                {
                                    if (!hunspell.Spell(result[i]))
                                    {
                                        result.RemoveAt(i);
                                    }
                                }
                            }

                            Methods.LogProceso("\n");
                            Methods.LogProceso(string.Format("Mean confidence: {0}", page.GetMeanConfidence()));
                            Methods.LogProceso(string.Format("Text (GetText): \r\n{0}", text));
                            Methods.LogProceso("Text (iterator):");

                            using (var iter = page.GetIterator())
                            {
                                iter.Begin();

                                do
                                {
                                    do
                                    {
                                        do
                                        {
                                            do
                                            {
                                                if (iter.IsAtBeginningOf(PageIteratorLevel.Block))
                                                {
                                                    Methods.LogProceso("<BLOCK>");
                                                }

                                                Methods.LogProceso(iter.GetText(PageIteratorLevel.Word));
                                                Methods.LogProceso(" ");

                                                if (iter.IsAtFinalOf(PageIteratorLevel.TextLine, PageIteratorLevel.Word))
                                                {
                                                    Methods.LogProceso("\n");
                                                }
                                            } while (iter.Next(PageIteratorLevel.TextLine, PageIteratorLevel.Word));

                                            if (iter.IsAtFinalOf(PageIteratorLevel.Para, PageIteratorLevel.TextLine))
                                            {
                                                Methods.LogProceso("\n");
                                            }
                                        } while (iter.Next(PageIteratorLevel.Para, PageIteratorLevel.TextLine));
                                    } while (iter.Next(PageIteratorLevel.Block, PageIteratorLevel.Para));
                                } while (iter.Next(PageIteratorLevel.Block));
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Trace.TraceError(e.ToString());
                Console.WriteLine("Unexpected Error: " + e.Message);
                Console.WriteLine("Details: ");
                Console.WriteLine(e.ToString());
            }
            //Console.Write("Press any key to continue . . . ");
            //Console.ReadKey(true);
        }

        //public void Login()
        //{
        //    try
        //    {

        //        if (Convert.ToBoolean(ConfigurationManager.AppSettings["DualValidation"].ToString().Trim()))
        //        {
        //            new ScreenPreviousLogin().DoActivities();
        //        }

        //        new ScreenLogin().DoActivities();
        //    }
        //    catch (Exception exc)
        //    {
        //        throw exc;
        //    }
        //}
    }
}
