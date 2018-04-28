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
using Ghostscript.NET.Rasterizer;

namespace To.Rpa.AppCTS.BL
{
    public class BLMain
    {
        private FileInfo[] OriginFiles { get; set; }
        private string OriginSharedDirectory { get; set; }
        private string OriginDirectory { get; set; }
        private string ScrapsDirectory { get; set; }
        private string TemplatesDirectory { get; set; }
        private string PdfsErrorDirectory { get; set; }
        private List<BECTSFormat> ApprovedFormats { get; set; }
        private List<BECTSFormat> DisapprovedFormats { get; set; }
        private List<BECTSImagePage> col2ImagesPagesTemp { get; set; }
        //private string col2FileNameTemp { get; set; }
        private bool foundHeader { get; set; }
        private string tesseractLanguage { get; set; }


        private UInt16 XScaling { get; set; }
        private UInt16 YScaling { get; set; }
        private UInt16 ScalingInitialPoint { get; set; }
        private UInt16 VerticalWidth { get; set; }
        private UInt16 VerticalHeight { get; set; }
        private string JarDirectory { get; set; }
        private string DllDirectory { get; set; }

        public BLMain()
        {
            OriginSharedDirectory = ConfigurationManager.AppSettings["OriginSharedDirectory"];
            OriginDirectory = ConfigurationManager.AppSettings["OriginDirectory"];
            ScrapsDirectory = ConfigurationManager.AppSettings["ScrapsDirectory"];
            TemplatesDirectory = ConfigurationManager.AppSettings["TemplatesDirectory"];
            tesseractLanguage = ConfigurationManager.AppSettings["TesseractLanguage"];
            PdfsErrorDirectory = ConfigurationManager.AppSettings["PdfsErrorDirectory"];
            XScaling = Convert.ToUInt16(ConfigurationManager.AppSettings["XScaling"]);
            YScaling = Convert.ToUInt16(ConfigurationManager.AppSettings["YScaling"]);
            ScalingInitialPoint = Convert.ToUInt16(ConfigurationManager.AppSettings["ScalingInitialPoint"]);
            VerticalWidth = Convert.ToUInt16(ConfigurationManager.AppSettings["VerticalWidth"]);
            VerticalHeight = Convert.ToUInt16(ConfigurationManager.AppSettings["VerticalHeight"]);
            JarDirectory = ConfigurationManager.AppSettings["JarDirectory"];
            DllDirectory = ConfigurationManager.AppSettings["DllDirectory"];

            //TODO, no debería ser lazy load... .
            ApprovedFormats = new List<BECTSFormat>();
            DisapprovedFormats = new List<BECTSFormat>();
            col2ImagesPagesTemp = new List<BECTSImagePage>();
        }

        public void DoActivities()
        {
            //GetPdfFromSharedDirectory();/*1*/
            //GetPdfFromMailBox(); MUEVE A \WORKSPACE\ORIGENES
            CreateFoldersPerCTSFormat();/*1*/
            //ExtractImagesFromPdf();
            //ImproveImagesFromPdf(); /*1*/
            GetApprovedFormats();/*1*/

            //GetScrapsHeaderDataImages();
            GetScrapsDataImages(); /*1 - en teoría si llego hasta acá, todo estará bien aunque no necesarimente copiado al 100%*/
            //GetScrapsDataImagesFull();
            //GetHeaderDataFromScraps();
            GetDataFromScraps();/*1*/
            //GetDataFullFromRaw();/*1*/
            //MovePdfNoProcess();/*1*/ //lvl 01
            //MoveScrapsErrorProcess();/*1*/
            //MoveScrapsCorrectProcess();/*1*/

            Console.WriteLine("Termino OK");
            Methods.LogProceso("Termino OK");
        }

        private void ExtractImagesFromPdf(String pathImage, String pathPdf)
        {
            /*try
            {
                System.Diagnostics.Process clientProcess = new Process();
                clientProcess.StartInfo.FileName = "java";
                clientProcess.StartInfo.Arguments = @"-jar " + " " + JarDirectory + " " + "\"" + pathPdf + "\"" + " " + pathImage + " " + DllDirectory;
                clientProcess.Start();
            }
            catch (Exception e)
            {
                e.GetBaseException();
            }*/

            //Get all the PDF files in the specified location
            //var pdfFiles = Directory.GetFiles(@"D:\PDFSplit\Example\outputFolder\", "*.pdf");

            //process each PDF file
            //foreach (var pdfFile in pdfFiles)
            //{
            //var fileName = Path.GetFileNameWithoutExtension(pathPdf);
            PdfToPng(pathImage, pathPdf);
            //}

        }

        private static void PdfToPng(string inputFile, string outputFileName)
        {
            var xDpi = 400; //set the x DPI
            var yDpi = 400; //set the y DPI
            var pageNumber = 1; // the pages in a PDF document

            
            try
            {
                using (var rasterizer = new GhostscriptRasterizer()) //create an instance for GhostscriptRasterizer
                {
                    
                    rasterizer.Open(outputFileName); //opens the PDF file for rasterizing

                    int pages = rasterizer.PageCount;
                    string outputPNGPath = "";
                    for (int i = 0; i < pages; i++)
                    {
                        outputPNGPath = Path.Combine(inputFile,"original", "base_" + inputFile.Substring(inputFile.Length - 4, 4) + "_" + i + ".png");
                        
                        var pdf2PNG = rasterizer.GetPage(xDpi, yDpi, i + 1);
                        pdf2PNG.Save(outputPNGPath, System.Drawing.Imaging.ImageFormat.Png);
                        Console.WriteLine("Saved " + outputPNGPath);
                    }

                }
            }
            catch (Exception exc)
            {
                Console.WriteLine(exc.Message);
            }
        }

        private void GetDataFullFromRaw()
        {
            List<string> data;
            foreach (BECTSFormat format in DisapprovedFormats)
            {
                foreach (BECTSImagePage imagePage in format.LstBECTSImagePages)
                {
                    data = GetData(ScrapsDirectory + format.FormatCode + @"\original\" + imagePage.CodePage);

                    if (data != null)
                    {
                        using (TextWriter tw = new StreamWriter(ScrapsDirectory + @"\" + format.FormatCode + @"\plantilla\" + Path.GetFileNameWithoutExtension(imagePage.CodePage) + ".txt"))
                        {
                            foreach (string s in data)
                                tw.WriteLine(s);
                        }
                    }
                }
            }
        }

        private void ImproveImagesFromPdf()
        {
            ReescalingImage();
        }

        private void ReescalingImage() //positive.
        {
            DirectoryInfo di = new DirectoryInfo(ScrapsDirectory);
            string directoryFullName = "";
            string directoryName = "";

            DirectoryInfo[] directories = di.GetDirectories();
            //FileInfo[] files = di.GetFiles();

            int totalDirectories = directories.Length;

            for (int i = 0; i < totalDirectories; i++)
            {
                //TODO: pensemos que vamos a reescalar solo las imagenes de data que vendrán solas sin nada más.
                directoryFullName = directories[i].FullName;
                directoryName = directories[i].Name;

                di = new DirectoryInfo(directoryFullName + @"\ORIGINAL\");

                FileInfo[] files = di.GetFiles("base_" + directoryName + "*");

                int totalArchivos = files.Length;

                for (int o = 0; o < totalArchivos; o++)
                {
                    Bitmap image = new Bitmap(files[o].FullName);
                    string eje = image.Width > image.Height ? "horizontal" : "vertical";

                    Bitmap newImage = null;
                    int yScaling = this.YScaling;
                    List<string> data = new List<string>();
                    int result = 0;

                    if (eje == "vertical")
                    {
                        do
                        {
                            newImage = Methods.cropAtRect(image, new Rectangle((image.Width / 2) - (this.XScaling / 2), (image.Height / 2) - (yScaling / 2), this.XScaling, 100));
                            newImage.Save(di.FullName + @"\Scaling.png");

                            data = GetData(di.FullName + @"\Scaling.png");

                            yScaling += 100;
                        } while (data.Count > 0);

                        result = (image.Height / 2) - (yScaling / 2) + this.YScaling;
                        result = result - this.ScalingInitialPoint;

                        newImage = Methods.cropAtRect(image, new Rectangle(result, result, image.Width - result, image.Height - result));
                        newImage.Save(di.FullName + @"\Scaling_final.png");

                        image.Dispose();
                        Size newSize = new Size(this.VerticalWidth, this.VerticalHeight);
                        image = new Bitmap(newImage, newSize);
                        image.Save(files[o].FullName);
                    }
                }

            }
        }

        private void MovePdfNoProcess()
        {
            DirectoryInfo diOrigin = new DirectoryInfo(OriginDirectory);
            DirectoryInfo diDestiny = new DirectoryInfo(PdfsErrorDirectory);

            //FileInfo[] files = OriginFiles;
            FileInfo[] files = diOrigin.GetFiles();

            int totalFiles = files.Length;

            for (int i = 0; i < totalFiles; i++)
            {
                File.Move(OriginDirectory + files[i].Name,
                    PdfsErrorDirectory + files[i].Name);
            }
        }

        private void MoveToUserFinal()
        {
            foreach (BECTSFormat format in DisapprovedFormats)
            {

            }
        }

        private void MoveScrapsErrorProcess()
        {
            foreach (BECTSFormat format in DisapprovedFormats)
            {

            }
        }

        private void MoveScrapsCorrectProcess()
        {
            foreach (BECTSFormat format in ApprovedFormats) { }
        }

        private void MoveBackupAll()
        {

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

                    ExtractImagesFromPdf(ScrapsDirectory + formatCode, OriginDirectory + files[i].Name);

                    File.Move(OriginDirectory + files[i].Name,
                        ScrapsDirectory + formatCode + @"\ORIGINAL\" + files[i].Name);
                }
            }
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
                    try
                    {
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
                    catch (Exception exc)
                    {
                        Console.WriteLine("algún archivo ya existe en la carpeta origenes con respecto al buzón");
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
                List<BECTSImagePage> lstBECTSImagePages = new List<BECTSImagePage>();
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
                            File.Copy(files[o].FullName, files[o].FullName.Replace(@"\ORIGINAL", ""), true);
                            attemps = Int32.Parse(ConfigurationManager.AppSettings["Attempts"]);
                        }
                        catch (Exception exc) { attemps++; }
                    }

                    for (int u = 0; u < 4; u++)
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
                                //TODO: JRDC, el tajo desde más arriba así sea header... .
                                newImage = Methods.cropAtRect(image, new Rectangle(3140, 2988, 700, image.Height - 2988));
                                //newImage = Methods.cropAtRect(image, new Rectangle(1570, 1494, 350, image.Height - 1494));
                            }
                            else
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(3140, 20, 700, image.Height - 20));
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
                                    lstBECTSImagePages.Add(new BECTSImagePage { CodePage = files[o].Name, isHeader = true });
                                    col2ImagesPagesTemp.Add(new BECTSImagePage
                                    {
                                        CodePage = files[o].Name,
                                        isHeader = true,
                                        ImagePagesScraps = new List<BECTSScrap>()
                                        {
                                            new BECTSScrap { ScrapPath = fileName }
                                        }
                                    });

                                    foundHeader = true;
                                }
                                else
                                {
                                    lstBECTSImagePages.Add(new BECTSImagePage { CodePage = files[o].Name });
                                    col2ImagesPagesTemp.Add(new BECTSImagePage
                                    {
                                        CodePage = files[o].Name,
                                        ImagePagesScraps = new List<BECTSScrap>()
                                        {
                                            new BECTSScrap { ScrapPath = fileName }
                                        }
                                    });
                                }
                                break;
                            }
                            else
                            {
                                ImageRotate(image, files[o].FullName.Replace(@"\ORIGINAL", ""));
                            }

                            //Me aseguro que termine la imagen en horizontal.
                            if (u == 3)
                            {
                                if (image.Width < image.Height)
                                {
                                    ImageRotate(image, files[o].FullName.Replace(@"\ORIGINAL", ""));
                                }

                            }
                        }

                    }
                    //falta donde guadan los diisprovaide
                    /*if (!foundHeader)
                                {
                                    lstBECTSImagePages.Add(new BECTSImagePage { CodePage = files[o].Name, isHeader = true });
                                    col2ImagesPagesTemp.Add(new BECTSImagePage
                                    {
                                        CodePage = files[o].Name,
                                        isHeader = true,
                                        ImagePagesScraps = new List<BECTSScrap>()
                                        {
                                            new BECTSScrap { ScrapPath = fileName }
                                        }
                                    });

                                    foundHeader = true;
                                }
                                else
                                {
                                    lstBECTSImagePages.Add(new BECTSImagePage { CodePage = files[o].Name });
                                    col2ImagesPagesTemp.Add(new BECTSImagePage
                                    {
                                        CodePage = files[o].Name,
                                        ImagePagesScraps = new List<BECTSScrap>()
                                        {
                                            new BECTSScrap { ScrapPath = fileName }
                                        }
                                    });
                                }
                                break;*/

                }

                if (approved)
                {
                    ApprovedFormats.Add(new BECTSFormat
                    {
                        FormatCode = int.Parse(directoryName),
                        LstBECTSImagePages = lstBECTSImagePages
                    });

                    approved = false;
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
                foreach (BECTSImagePage imagePages in formats.LstBECTSImagePages)
                {
                    string fileName = ScrapsDirectory + formats.FormatCode + @"\" + imagePages.CodePage;
                    using (Bitmap image = new Bitmap(fileName))
                    {
                        //Solo se procesan las imágenes en horizontal
                        if (image.Width > image.Height)
                        {
                            Bitmap newImage = null;
                            if (imagePages.isHeader)
                            {
                                //TODO: debería dejar todos en horizontal.
                                //RUC
                                newImage = Methods.cropAtRect(image, new Rectangle(2200, 200, 1500, 700));
                                fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_Header_RUC" + Path.GetExtension(imagePages.CodePage);
                                newImage.Save(fileName);
                                imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = true, ScrapPath = fileName });

                                //CODIGO_LISTADO
                                newImage = Methods.cropAtRect(image, new Rectangle(7480, 200, image.Width - 7480, 700));
                                fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_Header_FormatCode" + Path.GetExtension(imagePages.CodePage);
                                newImage.Save(fileName);
                                imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = true, ScrapPath = fileName });

                                //TODO:PUEDE OBTENERSE EL PERIODO DE OTRA MANERA.
                                //PERIODO
                                newImage = Methods.cropAtRect(image, new Rectangle(2200, 340, 1500, 700));
                                fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_Header_Period" + Path.GetExtension(imagePages.CodePage);
                                newImage.Save(fileName);
                                imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = true, ScrapPath = fileName });

                                //AÑO
                                newImage = Methods.cropAtRect(image, new Rectangle(4880, 340, 1500, 700));
                                fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_Header_Year" + Path.GetExtension(imagePages.CodePage);
                                newImage.Save(fileName);
                                imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = true, ScrapPath = fileName });

                                //MONEDA ABONO
                                newImage = Methods.cropAtRect(image, new Rectangle(7480, 200, image.Width - 7480, 860));
                                fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_Header_Currency" + Path.GetExtension(imagePages.CodePage);
                                newImage.Save(fileName);
                                imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = true, ScrapPath = fileName });
                            }

                            //COLUMNA 2 - NOMBRES Y APELLIDOS
                            if (imagePages.isHeader)
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(900, 2988, 1700, image.Height - 2988));
                            }
                            else
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(900, 20, 1700, image.Height - 20));
                            }
                            fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_2" + Path.GetExtension(imagePages.CodePage);
                            newImage.Save(fileName);
                            imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = false, ScrapPath = fileName });

                            //COLUMNA 3 - TIPO DOCUMENTO
                            if (imagePages.isHeader)
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(2640, 2988, 500, image.Height - 2988));
                            }
                            else
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(2640, 20, 500, image.Height - 20));
                            }
                            fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_3" + Path.GetExtension(imagePages.CodePage);
                            newImage.Save(fileName);
                            imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = false, ScrapPath = fileName });

                            //COLUMNA 4 - NRO DOCUMENTO
                            //imagePages.ImagePagesScraps.Add(col2FileNameTemp);
                            imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = false, ScrapPath = col2ImagesPagesTemp.Where(x => x.CodePage == imagePages.CodePage).Select(x => x.ImagePagesScraps[0].ScrapPath).First() });

                            //newImage = Methods.cropAtRect(image, new Rectangle(1570, 1494, 350, image.Height - 1494));
                            //fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_3" + Path.GetExtension(imagePages.CodePage);
                            //newImage.Save(fileName);
                            //imagePages.ImagePagesScraps.Add(fileName);

                            //COLUMNA 5 - NRO CUENTA
                            if (imagePages.isHeader)
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(3840, 2988, 800, image.Height - 2988));
                            }
                            else
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(3840, 20, 800, image.Height - 20));
                            }
                            fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_5" + Path.GetExtension(imagePages.CodePage);
                            newImage.Save(fileName);
                            imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = false, ScrapPath = fileName });

                            //COLUMNA 6 - CUENTA > MONEDA
                            if (imagePages.isHeader)
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(4840, 2988, 440, image.Height - 2988));
                            }
                            else
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(4840, 20, 440, image.Height - 20));
                            }
                            fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_6" + Path.GetExtension(imagePages.CodePage);
                            newImage.Save(fileName);
                            imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = false, ScrapPath = fileName });

                            //COLUMNA 7 - ABONO > MONTO
                            if (imagePages.isHeader)
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(5320, 2988, 760, image.Height - 2988));
                            }
                            else
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(5320, 20, 760, image.Height - 20));
                            }
                            fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_7" + Path.GetExtension(imagePages.CodePage);
                            newImage.Save(fileName);
                            imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = false, ScrapPath = fileName });

                            //COLUMNA 8 - ABONO > MONEDA
                            if (imagePages.isHeader)
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(6080, 2988, 440, image.Height - 2988));
                            }
                            else
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(6080, 20, 440, image.Height - 20));
                            }
                            fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_8" + Path.GetExtension(imagePages.CodePage);
                            newImage.Save(fileName);
                            imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = false, ScrapPath = fileName });

                            //COLUMNA 9 - 4 REMUNERACIONES > MONTO
                            if (imagePages.isHeader)
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(6720, 2988, 960, image.Height - 2988));
                            }
                            else
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(6720, 20, 960, image.Height - 20));
                            }
                            fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_9" + Path.GetExtension(imagePages.CodePage);
                            newImage.Save(fileName);
                            imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = false, ScrapPath = fileName });

                            //COLUMNA 10 - 4 REMUNERACIONES > MONEDA
                            if (imagePages.isHeader)
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(7680, 2988, 440, image.Height - 2988));
                            }
                            else
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(7680, 20, 440, image.Height - 20));
                            }
                            fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_10" + Path.GetExtension(imagePages.CodePage);
                            newImage.Save(fileName);
                            imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = false, ScrapPath = fileName });

                            //imagePages.ImagePagesScraps = imagePages.ImagePagesScraps.OrderBy(x => x).ToList();
                        }
                    }
                }
            }
        }

        //private void GetScrapsHeaderDataImages()
        //{
        //    foreach (BECTSFormat formats in ApprovedFormats)
        //    {
        //        foreach (BECTSImagePage imagePages in formats.LstBECTSImagePages)
        //        {
        //            string fileName = ScrapsDirectory + formats.FormatCode + @"\" + imagePages.CodePage;
        //            using (Bitmap image = new Bitmap(fileName))
        //            {

        //                Bitmap newImage = null;
        //                if (imagePages.isHeader)
        //                {
        //                    //RUC
        //                    newImage = Methods.cropAtRect(image, new Rectangle(1100, 100, 750, 350));
        //                    fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_Header_RUC" + Path.GetExtension(imagePages.CodePage);
        //                    newImage.Save(fileName);
        //                    imagePages.ImagePagesScraps.Add(true, fileName);

        //                    //CODIGO_LISTADO
        //                    newImage = Methods.cropAtRect(image, new Rectangle(3740, 100, image.Width - 3740, 350));
        //                    fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_Header_FormatCode" + Path.GetExtension(imagePages.CodePage);
        //                    newImage.Save(fileName);
        //                    imagePages.ImagePagesScraps.Add(true, fileName);

        //                    //TODO:PUEDE OBTENERSE EL PERIODO DE OTRA MANERA.
        //                    //PERIODO
        //                    newImage = Methods.cropAtRect(image, new Rectangle(1100, 170, 750, 350));
        //                    fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_Header_Period" + Path.GetExtension(imagePages.CodePage);
        //                    newImage.Save(fileName);
        //                    imagePages.ImagePagesScraps.Add(true, fileName);

        //                    //AÑO
        //                    newImage = Methods.cropAtRect(image, new Rectangle(2440, 170, 750, 350));
        //                    fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_Header_Year" + Path.GetExtension(imagePages.CodePage);
        //                    newImage.Save(fileName);
        //                    imagePages.ImagePagesScraps.Add(true, fileName);

        //                    //MONEDA ABONO
        //                    newImage = Methods.cropAtRect(image, new Rectangle(3740, 100, image.Width - 3740, 430));
        //                    fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_Header_Currency" + Path.GetExtension(imagePages.CodePage);
        //                    newImage.Save(fileName);
        //                    imagePages.ImagePagesScraps.Add(true, fileName);
        //                }

        //            }
        //        }
        //    }
        //}

        private void ImageRotate(Image image, string fileName)
        {
            image.RotateFlip(RotateFlipType.Rotate90FlipXY);
            image.Save(fileName);
        }

        //private void GetHeaderDataFromScraps()
        //{
        //    foreach (BECTSFormat formats in ApprovedFormats)
        //    {
        //        //COPIO PLANTILLA CON NOMBRE DE FORMATO.
        //        string filename = ScrapsDirectory + formats.FormatCode + @"\plantilla\" + "Plantilla abono CTS" + " - " + formats.FormatCode + ".xlsx";

        //        //TODO: validar si el archivo existe y marcarlo en el log.
        //        File.Copy(TemplatesDirectory + "Plantilla abono CTS.xlsx", filename, true);
        //        using (SpreadsheetDocument document = SpreadsheetDocument.Open(filename, true))
        //        {
        //            //TODO: Validar si el header no existe no se procesa nada... a menos que se puedan traer los datos de algún lado adicional...
        //            string t = formats.LstBECTSImagePages[0].CodePage + "_0_Header_RUC";
        //            List<string> dataCol2 = GetData(t);

        //            t = formats.LstBECTSImagePages[0].CodePage + "_0_Header_FormatCode";
        //            dataCol2 = GetData(t);

        //            t = formats.LstBECTSImagePages[0].CodePage + "_0_Header_Period";
        //            dataCol2 = GetData(t);
        //            t = formats.LstBECTSImagePages[0].CodePage + "_0_Header_Year";
        //            dataCol2 = GetData(t);
        //            t = formats.LstBECTSImagePages[0].CodePage + "_0_Header_Currency";
        //            dataCol2 = GetData(t);
        //        }
        //    }
        //}

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

                    foreach (BECTSImagePage imagePages in formats.LstBECTSImagePages)
                    {
                        int initialCol = 2;
                        string col = "";
                        //TODO: el salto de lineas en blanco debe ser parámetro, o hago método que evalue q tanto se parece la cadena al nombre que devuelve la búsqueda por dni en la reniec. o ignoro la columna nombres y apellidos desde el pdf y la pinto desde reniec.
                        //le agregamos más 2  espacios en blanco en caso se salte uno o máximo dos por error.
                        lastRowCountImagePage += lastRowCountImagePageTemp;
                        foreach (BECTSScrap scrap in imagePages.ImagePagesScraps)
                        {
                            if (!scrap.isHeaderValue)
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
                                List<string> dataCol2 = GetData(scrap.ScrapPath);
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
                            else
                            {
                                //TODO: Validar si el header no existe no se procesa nada... a menos que se puedan traer los datos de algún lado adicional...
                                List<string> dataCol2 = null;

                                //TODO: VALIDA RUC
                                switch (imagePages.ImagePagesScraps.IndexOf(scrap))
                                {
                                    case 0:
                                        dataCol2 = GetData(scrap.ScrapPath);
                                        UpdateCell(document, repairRUC(dataCol2), 5, "D"); break;
                                    case 1:
                                        //dataCol2 = GetData(scrap.ScrapPath);
                                        UpdateCell(document, formats.FormatCode.ToString(), 5, "K"); break;
                                    case 2:
                                        dataCol2 = GetData(scrap.ScrapPath);
                                        UpdateCell(document, repairPeriod(dataCol2), 7, "D"); break;
                                    case 3:
                                        dataCol2 = GetData(scrap.ScrapPath);
                                        UpdateCell(document, repairYear(dataCol2), 7, "H"); break;
                                    case 4:
                                        dataCol2 = GetData(scrap.ScrapPath);
                                        UpdateCell(document, /*repairCurrency(dataCol2)*/ "DOL", 7, "K"); break;

                                }
                            }
                        }
                    }
                }
            }
        }
        private string repairYear(List<string> lst)
        {
            string year = "";
            DateTime yearTemp;
            foreach (string str in lst)
            {
                if (str.Trim() == "")
                    continue;
                year = Regex.Replace(str, @"[^\d]", String.Empty);

                if (Regex.IsMatch(year, @"\d{4}"))
                {
                    if (DateTime.TryParse("01/01/" + year, out yearTemp))
                        return year;
                }
            }
            return "";
            //TODO: grabar en log si el ruc está en vacío.
        }

        private string repairCurrency(List<string> lst)
        {
            string currency = "";
            List<string> currencies = new List<string>() { "SOL", "DOL" };

            foreach (string str in lst)
            {
                if (str.Trim() == "")
                    continue;
                currency = Regex.Replace(str, @"[^\p{L}]", String.Empty);

                if (currencies.Contains(currency))
                    return currency;
            }
            return "";
            //TODO: grabar en log si el ruc está en vacío.
        }

        private string repairPeriod(List<string> lst)
        {
            List<string> periods = new List<string>() { "MAY", "NOV" };
            string period = "";
            foreach (string str in lst)
            {
                if (str.Trim() == "")
                    continue;
                period = Regex.Replace(str, @"[^\p{L}]", String.Empty);

                if (periods.Contains(period))
                    return period;
            }
            return "";
        }

        private string repairRUC(List<string> lst)
        {
            string ruc = "";
            foreach (string str in lst)
            {
                if (str.Trim() == "")
                    continue;
                ruc = Regex.Replace(str, @"[^\d]", String.Empty).PadLeft(11, '0');

                if (Regex.IsMatch(ruc, @"\d{11}"))
                    return ruc;
            }
            return "";
            //TODO: grabar en log si el ruc está en vacío.
        }

        private string repairID(string names, string ID)
        {
            //TODO, poner rojo si no cumple.
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
                using (var engine = new TesseractEngine(@"./Dictionaries/Tesseract/eng", "eng", EngineMode.Default))
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

            if (lstIDs.Count > 0)
            {
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

                if (accepted >= minimunAccepted)
                {
                    //TODO: guardar lo le´´ido para optimziar.
                    return true;
                }
                return false;
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
/*
el log con un orden...
integrar método de validación ibs
que muevan las plantillas finanles de cada format code a carpeta final.
que las que no se aprueban pasen a error

probar

identificar errores.

obtener datos de la cabecera.

hacer bucle infito en el main que pare con un valor en un archivo de texto.

revisar todos y realizar lo que se indica.


corregir las dimensiones de una manera dinámica

incluir parte de felix.*/
