using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Ghostscript.NET;
using Ghostscript.NET.Rasterizer;
using IBM.Data.DB2.iSeries;
using NHunspell;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using Tesseract;
using To.Rpa.AppCTS.BE;
using To.Rpa.Util;
using A = DocumentFormat.OpenXml.OpenXmlAttribute;
using E = DocumentFormat.OpenXml.OpenXmlElement;
using S = DocumentFormat.OpenXml.Spreadsheet.Sheets;

namespace To.Rpa.AppCTS.BL
{
    public class BLMain
    {
        private string OriginSharedDirectory { get; set; }
        private string OriginDirectory { get; set; }
        private string ScrapsDirectory { get; set; }
        private string ManualScrapsDirectory { get; set; }
        private string TemplatesDirectory { get; set; }
        private string WrongPdfsDirectory { get; set; }
        private string FinalFormatDirectory { get; set; }
        private string GSDllDirectory { get; set; }
        private string BackupDirectory { get; set; }

        private List<BECTSFormat> ApprovedFormats { get; set; }
        private List<BECTSFormat> DisapprovedFormats { get; set; }
        private List<BECTSImagePage> Col2ImagesPagesTemp { get; set; }
        private bool FoundHeader { get; set; }
        private string TesseractLanguage { get; set; }

        private UInt16 XScaling { get; set; }
        private UInt16 YScaling { get; set; }
        private UInt16 ScalingInitialPoint { get; set; }
        private UInt16 VerticalWidth { get; set; }
        private UInt16 VerticalHeight { get; set; }
        private int XDpi { get; set; }
        private int YDpi { get; set; }
        private string JarDirectory { get; set; }
        private UInt16 Attempts { get; set; }

        private bool StopAuto { get; set; }
        private bool RestAuto { get; set; }
        private bool StopManual { get; set; }
        private bool RestManual { get; set; }
        private int Sleep { get; set; }

        public BLMain()
        {
            OriginSharedDirectory = ConfigurationManager.AppSettings["OriginSharedDirectory"];
            OriginDirectory = ConfigurationManager.AppSettings["OriginDirectory"];
            ScrapsDirectory = ConfigurationManager.AppSettings["ScrapsDirectory"];
            ManualScrapsDirectory = ConfigurationManager.AppSettings["ManualScrapsDirectory"];
            TemplatesDirectory = ConfigurationManager.AppSettings["TemplatesDirectory"];
            TesseractLanguage = ConfigurationManager.AppSettings["TesseractLanguage"];
            WrongPdfsDirectory = ConfigurationManager.AppSettings["WrongPdfsDirectory"];
            FinalFormatDirectory = ConfigurationManager.AppSettings["FinalFormatDirectory"];
            XScaling = Convert.ToUInt16(ConfigurationManager.AppSettings["XScaling"]);
            YScaling = Convert.ToUInt16(ConfigurationManager.AppSettings["YScaling"]);
            ScalingInitialPoint = Convert.ToUInt16(ConfigurationManager.AppSettings["ScalingInitialPoint"]);
            VerticalWidth = Convert.ToUInt16(ConfigurationManager.AppSettings["VerticalWidth"]);
            VerticalHeight = Convert.ToUInt16(ConfigurationManager.AppSettings["VerticalHeight"]);
            XDpi = Convert.ToUInt16(ConfigurationManager.AppSettings["XDpi"]);
            YDpi = Convert.ToUInt16(ConfigurationManager.AppSettings["YDpi"]);
            Attempts = Convert.ToUInt16(ConfigurationManager.AppSettings["Attempts"]);
            GSDllDirectory = ConfigurationManager.AppSettings["GSDllDirectory"];
            BackupDirectory = ConfigurationManager.AppSettings["BackupDirectory"];

            StopAuto = Convert.ToBoolean(ConfigurationManager.AppSettings["StopAuto"]);
            RestAuto = Convert.ToBoolean(ConfigurationManager.AppSettings["RestAuto"]);
            StopManual = Convert.ToBoolean(ConfigurationManager.AppSettings["StopManual"]);
            RestManual = Convert.ToBoolean(ConfigurationManager.AppSettings["RestManual"]);
            Sleep = Convert.ToInt32(ConfigurationManager.AppSettings["Sleep"]);

            //TODO, no debería ser lazy load... .
            //ApprovedFormatsByName = new List<string>();
            ApprovedFormats = new List<BECTSFormat>();
            DisapprovedFormats = new List<BECTSFormat>();
            Col2ImagesPagesTemp = new List<BECTSImagePage>();

        }

        public void DoActivities()
        {
            Log("Inicia Karenx 0.9");

            //TODO: revisar el uso de Thread.CurrentThread.IsBackground = true; y de task async etc.
            Thread automaticProcess = new Thread(() =>
            {
                do
                {

                    ////config.AppSettings.SectionInformation.ForceSave = true;
                    //ConfigurationManager.RefreshSection("appSettings");
                    //System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                    ////System.Configuration.Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                    //TODO: para que se pueda actualizar los valores.
                    StopAuto = Convert.ToBoolean(ConfigurationManager.AppSettings["StopAuto"]);
                    RestAuto = Convert.ToBoolean(ConfigurationManager.AppSettings["RestAuto"]);
                    Sleep = Convert.ToInt32(ConfigurationManager.AppSettings["Sleep"]);

                    if (!RestAuto)
                    {
                        //Thread.CurrentThread.IsBackground = true;
                        GetFormatsFromSharedDirectory();/*1*/
                                                        //GetPdfFromMailBox();
                        CreateFoldersByEachFormat();/*1*/
                        ExtractImagesFromFormats();/*1*/
                                                   //ImproveImagesFromPdf(); /*1*/
                        GetApprovedFormats();/*1*/

                        CreateFoldersByEachDisapprovedFormat();/*1*/

                        //GetScrapsHeaderDataImages();
                        GetScrapsDataImages(); /*1 - en teoría si llego hasta acá, todo estará bien aunque no necesarimente copiado al 100%*/
                                               //GetScrapsDataImagesFull();
                                               //GetHeaderDataFromScraps();
                        GetDataFromScraps();/*1*/
                                            //GetDataManualFromScraps();/*1*/
                                            //GetDataFullFromRaw();/*1*/
                        MoveFinalFormats(); /*1*/
                        MoveFormatsNoProcess();/*1*/ //lvl 01
                        //MoveScrapsErrorProcess();
                        BackupScrapsCorrectProcess();/*1*/

                        ApprovedFormats.Clear();
                        DisapprovedFormats.Clear();
                    }
                    else
                    {
                        Log("Descanso total de proceso automático... .");
                        Thread.Sleep(Sleep);
                    }
                }
                while (!StopAuto);

                Log("Detengo proceso automático de forma normal");
            });

            Thread manualProcess = new Thread(() =>
            {
                do
                {
                    StopManual = Convert.ToBoolean(ConfigurationManager.AppSettings["StopManual"]);
                    RestManual = Convert.ToBoolean(ConfigurationManager.AppSettings["RestManual"]);
                    Sleep = Convert.ToInt32(ConfigurationManager.AppSettings["Sleep"]);

                    if (!RestManual)
                    {
                        Thread.CurrentThread.IsBackground = true;
                        /* run your code here */
                        Log("Inicio atención de proceso manual");
                        GetDataManualFromScraps();
                        MoveManualFinalFormats();
                        BackupManualScrapsCorrectProcess();
                        Log("Termino atención de proceso manual");
                        Log("Descanso por " + (Sleep / 1000) + " segundos");
                        Thread.Sleep(Sleep);
                    }
                    else
                    {
                        Log("Descanso total de proceso manual... .");
                        Thread.Sleep(Sleep);
                    }
                }
                while (!StopManual);

                Log("Detengo proceso manual de forma normal");
            });

            automaticProcess.Start();
            manualProcess.Start();

            Log("Rutina de inicialización Karenx 0.9 completada");
        }

        private void CreateFoldersByEachDisapprovedFormat()
        {
            foreach (BECTSFormat format in DisapprovedFormats)
            {
                Log("#4.1| Inicia rutina #4.1 Crea carpetas por cada formato desaprobado dentro de la ruta: " + ManualScrapsDirectory);
                try
                {
                    DirectoryInfo diScraps = new DirectoryInfo(ScrapsDirectory + format.FormatCode);
                    DirectoryInfo diManual = new DirectoryInfo(ManualScrapsDirectory);

                    if (diManual.GetDirectories(format.FormatCode + "*").Length == 0)
                    {
                        diScraps.MoveTo(ManualScrapsDirectory + format.FormatCode);
                        Log("#4.1| Se movió hacia " + ManualScrapsDirectory + format.FormatCode);
                    }
                    else
                    {
                        Log("#4.1| Ya existe el directorio " + ManualScrapsDirectory + format.FormatCode);
                    }

                    //diManual.GetDirectories("PLANTILLA")[0].Delete(true);
                }
                catch (Exception exc)
                {
                    Methods.LogProceso("#4.1| ERROR: " + exc.ToString());
                }
            }

            Log("#4.1| Termina rutina #4.1");
        }

        private void MoveFinalFormats()
        {
            Log("#7| Inicia rutina #7");
            try
            {
                foreach (BECTSFormat format in ApprovedFormats)
                {
                    //   d: \Users\juarui\Source\Repos\RPA\To.Rpa.AppCTS\CTS\WORKSPACE\RETAZOS\4023\PLANTILLA

                    DirectoryInfo finalFormatDirectory = new DirectoryInfo(Path.Combine(ScrapsDirectory, format.FormatCode.ToString()));

                    DirectoryInfo userFinal = new DirectoryInfo(Path.Combine(ScrapsDirectory, format.FormatCode.ToString()));

                    DirectoryInfo[] directories = userFinal.GetDirectories(format.FormatCode + "*");
                     
                    if (directories.Length > 0)
                    {
                        finalFormatDirectory.MoveTo(FinalFormatDirectory + "_" + directories.Length);
                    }
                    else
                    {
                        finalFormatDirectory.MoveTo(FinalFormatDirectory);
                    }
                    /*DirectoryInfo finalFormat = new DirectoryInfo(Path.Combine(ScrapsDirectory, format.FormatCode.ToString(), "PLANTILLA"));

                    FileInfo[] files = finalFormat.GetFiles("*.xls*");


                    foreach (FileInfo file in files)
                    {
                        //File.Move(file.FullName, FinalFormatDirectory + file.Name);
                        //TODO:cambiar por move más length...
                        File.Copy(file.FullName, FinalFormatDirectory + file.Name, true);
                    }*/
                }
            }
            catch (Exception exc)
            {
                Log("#7| " + exc.ToString());
            }
            Log("#7| Termina rutina #7");
        }

        private void MoveManualFinalFormats()
        {
            Log("#x| Inicia rutina #7");
            try
            {
                foreach (BECTSFormat format in DisapprovedFormats)
                {
                    if (format.Done)
                    {
                        //   d: \Users\juarui\Source\Repos\RPA\To.Rpa.AppCTS\CTS\WORKSPACE\RETAZOS\4023\PLANTILLA
                        DirectoryInfo finalFormat = new DirectoryInfo(Path.Combine(ManualScrapsDirectory, format.FormatCode.ToString(), "PLANTILLA"));

                        FileInfo[] files = finalFormat.GetFiles("*.xls*");

                        foreach (FileInfo file in files)
                        {
                            //File.Move(file.FullName, FinalFormatDirectory + file.Name);
                            //TODO:cambiar por move más length...
                            File.Copy(file.FullName, FinalFormatDirectory + file.Name, true);
                        }
                    }
                }


            }
            catch (Exception exc)
            {
                Log("#x| " + exc.ToString());
            }
            Log("#x| Termina rutina");
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

        private void MoveFormatsNoProcess()
        {
            DirectoryInfo diOrigin = new DirectoryInfo(OriginDirectory);
            DirectoryInfo diDestiny = new DirectoryInfo(WrongPdfsDirectory);

            //FileInfo[] files = OriginFiles;
            FileInfo[] files = diOrigin.GetFiles();

            int totalFiles = files.Length;

            for (int i = 0; i < totalFiles; i++)
            {
                if (!File.Exists(WrongPdfsDirectory + files[i].Name))
                {
                    File.Move(OriginDirectory + files[i].Name,
                        WrongPdfsDirectory + files[i].Name);
                }
            }
        }

        private void MoveScrapsErrorProcess()
        {
            foreach (BECTSFormat format in DisapprovedFormats)
            {

            }
        }

        private void BackupScrapsCorrectProcess()
        {
            Log("#8| Inicia rutina");
            try
            {
                foreach (BECTSFormat format in ApprovedFormats)
                {
                    if (format.Done)
                    {
                        //   d: \Users\juarui\Source\Repos\RPA\To.Rpa.AppCTS\CTS\WORKSPACE\RETAZOS\4023\PLANTILLA
                        DirectoryInfo finalFormat = new DirectoryInfo(Path.Combine(ScrapsDirectory, format.FormatCode.ToString()));
                        finalFormat.MoveTo(BackupDirectory + format.FormatCode.ToString());
                    }
                }
            }
            catch (Exception exc)
            {
                Log("#8| ERROR: " + exc.ToString());
            }
            Log("#8| Termina rutina");
        }

        private void BackupManualScrapsCorrectProcess()
        {
            Log("#x2| Inicia rutina");
            try
            {
                foreach (BECTSFormat format in DisapprovedFormats)
                {
                    if (format.Done)
                    {
                        //   d: \Users\juarui\Source\Repos\RPA\To.Rpa.AppCTS\CTS\WORKSPACE\RETAZOS\4023\PLANTILLA
                        DirectoryInfo finalFormat = new DirectoryInfo(Path.Combine(ManualScrapsDirectory, format.FormatCode.ToString() + "_MANUAL"));
                        finalFormat.MoveTo(BackupDirectory + format.FormatCode.ToString());
                    }
                }
            }
            catch (Exception exc)
            {
                Log("#x2| ERROR: " + exc.ToString());
            }
            Log("#x2| Termina rutina");
        }

        private void MoveBackupAll()
        {

        }

        //TODO: por que los thumbnails de windows muestran imagenes que no corresponden a la realidad??
        /// <summary>
        /// Obtiene el código de formato, también llamado código de listado de los nombres de los formatos pdf, que terminen 
        /// en un guión más un número, caso contrario devuelve 0
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private int getFormatCode(string fileName)
        {
            //el formato es un guión (-), seguido de un número, se podría hacer más inteligente.
            string[] scraps = fileName.Split('-');

            int formatCode;
            int.TryParse(Path.GetFileNameWithoutExtension(scraps[scraps.Length - 1].Trim()), out formatCode);
            return formatCode;
        }

        /// <summary>
        /// 
        /// </summary>
        private void CreateFoldersByEachFormat()
        {
            Log("#2| Inicia rutina #2 Crea carpetas por cada formato a procesar de la carpeta: " + OriginDirectory + " dentro de la ruta: " + ScrapsDirectory);
            try
            {
                //TODO se puede cambiar por la lista de la variable de clasee...
                DirectoryInfo diOrigin = new DirectoryInfo(OriginDirectory);
                DirectoryInfo diScraps = new DirectoryInfo(ScrapsDirectory);
                DirectoryInfo diWrongPdfs = new DirectoryInfo(WrongPdfsDirectory);
                DirectoryInfo diScrapsChild = null;

                FileInfo[] originFiles = diOrigin.GetFiles();

                for (int i = 0; i < originFiles.Length; i++)
                {
                    int formatCode = getFormatCode(originFiles[i].Name);

                    if (formatCode != 0)
                    {
                        string parcialPath = Path.Combine(ScrapsDirectory, formatCode.ToString());

                        Log("#2| Formato a procesar " + formatCode);
                        //TODO: Checar si es que hay carpeta repetido cuidado se sobreescribirá..
                        diScrapsChild = new DirectoryInfo(parcialPath + @"\");
                        diScrapsChild.Create();
                        diScrapsChild.Refresh();

                        diScrapsChild = new DirectoryInfo(parcialPath + @"\ORIGINAL\");
                        diScrapsChild.Create();
                        diScrapsChild.Refresh();

                        diScrapsChild = new DirectoryInfo(parcialPath + @"\PLANTILLA\");
                        diScrapsChild.Create();
                        diScrapsChild.Refresh();

                        File.Move(OriginDirectory + originFiles[i].Name, parcialPath + @"\ORIGINAL\" + originFiles[i].Name);
                        Log("#2| Se movió archivo: " + originFiles[i].Name + " hacia: " + parcialPath + @"\ORIGINAL\" + originFiles[i].Name);

                        //ApprovedFormatsByName.Add(parcialPath + @"\ORIGINAL\" + originFiles[i].Name);
                    }
                    else
                    {
                        Log("#2| El archivo " + originFiles[i].Name + " no tiene el nombre de archivo en el formato correcto.");
                        File.Move(OriginDirectory + originFiles[i].Name, WrongPdfsDirectory + originFiles[i].Name);
                        Log("#2| Se movió archivo: " + originFiles[i].Name + " hacia: " + WrongPdfsDirectory + originFiles[i].Name);
                    }
                }
            }
            catch (Exception exc)
            {
                Methods.LogProceso("#2| ERROR: " + exc.ToString());
            }
            Log("#2| Termina rutina #2");
        }

        private void ExtractImagesFromFormats()
        {
            Log("#3| Inicia rutina #3 Extrae imágenes de formato pdf en la misma ruta por formato " + ScrapsDirectory);
            try
            {
                //TODO se puede cambiar por la lista de la variable de clasee...
                DirectoryInfo diScraps = new DirectoryInfo(ScrapsDirectory);
                DirectoryInfo[] directories = diScraps.GetDirectories();

                for (int i = 0; i < directories.Length; i++)
                {

                    int format = getFormatCode(directories[i].Name);
                    DirectoryInfo diChildScraps = new DirectoryInfo(ScrapsDirectory + format + @"\ORIGINAL\");

                    Log("#3| Formato a procesar " + format);

                    string imagePath = directories[i].FullName + @"\ORIGINAL\base_" + format + "_" + "{0}.png";

                    PdfToPng(diChildScraps.GetFiles("*.pdf")[0].FullName, imagePath, 0, XDpi, YDpi);
                }
            }
            catch (Exception exc)
            {
                Log("#3| ERROR: " + exc.ToString());
            }
            Log("#3| Termina rutina #3");
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="inputFilePath"></param>
        /// <param name="outPutFilePath"></param>
        /// <param name="pageNumber">Si se manda el número de página 0, se entiende que se desean extraer todas las páginas del pdf.</param>
        /// <param name="xDpi"></param>
        /// <param name="yDpi"></param>
        private void PdfToPng(string inputFilePath, string outPutFilePath, int pageNumber, int xDpi, int yDpi)
        {
            try
            {
                GhostscriptVersionInfo gvi = new GhostscriptVersionInfo(GSDllDirectory);

                using (var rasterizer = new GhostscriptRasterizer()) //create an instance for GhostscriptRasterizer
                {
                    rasterizer.Open(inputFilePath, gvi, true); //opens the PDF file for rasterizing
                    if (pageNumber == 0)
                    {
                        for (int i = 0; i < rasterizer.PageCount; i++)
                        {
                            var image = rasterizer.GetPage(xDpi, yDpi, i + 1);
                            string imagePath = string.Format(outPutFilePath, i);

                            image.Save(imagePath, System.Drawing.Imaging.ImageFormat.Png);
                            Log("Se extrae imagen de página " + (i + 1) + " de '" + Path.GetFileName(inputFilePath) + "' hacia '" + Path.GetFileName(imagePath) + "'");
                        }
                    }
                    else
                    {
                        var image = rasterizer.GetPage(xDpi, yDpi, pageNumber);
                        image.Save(string.Format(outPutFilePath, pageNumber), System.Drawing.Imaging.ImageFormat.Png);
                        Log("Se extrae imagen " + pageNumber + " de '" + Path.GetFileName(inputFilePath) + "' hacia '" + Path.GetFileName(outPutFilePath) + "'");
                    }

                }
            }
            catch (Exception exc)
            {
                Log(exc.ToString());
            }
        }


        private void GetPdfFromMailBox()
        {

        }
        //TODO guardar nombre de los archivos en la primera lectura sin criterios ... para luego moverlos... .,los que
        //quedan es porque no se procesaron.

        /// <summary>
        /// Obtiene los archivos pdf del directorio compartido OriginSharedDirectory path
        /// </summary>
        private void GetFormatsFromSharedDirectory()
        {
            Log("#1| Inicia rutina #1 Obtiene archivos origen de directorio compartido: " + OriginSharedDirectory);
            try
            {
                DirectoryInfo diOrigin = new DirectoryInfo(OriginSharedDirectory);
                FileInfo[] originFiles = diOrigin.GetFiles();

                for (int i = 0; i < originFiles.Length; i++)
                {
                    DirectoryInfo diDestionation = new DirectoryInfo(OriginDirectory);

                    FileInfo[] destinationFiles = diDestionation.GetFiles(originFiles[i].Name);

                    int length = destinationFiles.Length;

                    if (destinationFiles.Length > 0)
                    {
                        File.Move(originFiles[i].FullName,
                            diDestionation.FullName + GetCustomFileName(originFiles[i].Name, length.ToString()));
                        Log("#1| Formato ya existente, se colocó correlativo");
                        Log("#1| Se movió archivo: " + originFiles[i].Name + " hacia: " + diDestionation.FullName + GetCustomFileName(originFiles[i].Name, length.ToString()));
                    }
                    else
                    {
                        File.Move(diOrigin.FullName + originFiles[i].Name, diDestionation.FullName + originFiles[i].Name);
                        Log("#1| Se movió archivo: " + originFiles[i].Name + " hacia: " + diDestionation.FullName + originFiles[i].Name);
                    }
                }
            }
            catch (Exception exc)
            {
                Log("#1| ERROR: " + exc.ToString());
            }

            Log("#1| Termina rutina #1");
        }

        private void Log(string msg)
        {
            Console.WriteLine(msg);
            Methods.LogProceso(msg);
        }

        private string GetCustomFileName(string fileName, string correlative)
        {
            return fileName.Replace(".", "_" + correlative + ".");
        }

        private void GetApprovedFormats()
        {
            Log("#4| Inicia rutina #4 Obtiene formatos aprobados de manera automática");
            DirectoryInfo di = new DirectoryInfo(ScrapsDirectory);
            string fileName = "";
            string directoryFullName = "";
            string directoryName = "";
            bool approved = false;
            List<string> lstPathDirectories = new List<string>();

            DirectoryInfo[] directories = di.GetDirectories();
            try
            {

                for (int i = 0; i < directories.Length; i++)
                {
                    List<BECTSImagePage> lstBECTSImagePages = new List<BECTSImagePage>();
                    FoundHeader = false;

                    directoryFullName = directories[i].FullName;
                    directoryName = directories[i].Name;

                    di = new DirectoryInfo(directoryFullName + @"\ORIGINAL\");
                    FileInfo[] files = di.GetFiles("base_" + directoryName + "*");

                    if (directoryName == "4038")
                    {

                    }

                    for (int o = 0; o < files.Length; o++)
                    {
                        UInt16 attempts = 0;
                        while (attempts < Attempts)
                        {
                            try
                            {
                                File.Copy(files[o].FullName, files[o].FullName.Replace(@"\ORIGINAL", ""), true);
                                Log("#4| Copia archivo " + files[o].FullName + " a nivel superior");
                                attempts = Attempts;
                            }
                            catch (Exception exc)
                            {
                                Log(exc.ToString());
                                attempts++;
                            }
                        }

                        for (int u = 0; u < 4; u++)
                        {
                            using (Bitmap image = new Bitmap(files[o].FullName.Replace(@"\ORIGINAL", "")))
                            {
                                Bitmap newImage = null;

                                //AHORA SE EVALÚA POR NRO DOCUMENTO //COLUMNA 4 - NRO DOCUMENTO
                                if (!FoundHeader)
                                {
                                    newImage = Methods.cropAtRect(image, new Rectangle(1570, 1494, 350, image.Height - 1494));
                                }
                                else
                                {
                                    newImage = Methods.cropAtRect(image, new Rectangle(1570, 20, 350, image.Height - 20));
                                }
                                fileName = files[o].DirectoryName.Replace(@"\ORIGINAL", "") + @"\" + Path.GetFileNameWithoutExtension(files[o].Name) + "_4" + Path.GetExtension(files[0].Name);

                                newImage.Save(fileName);
                                //Methods.LogProceso("PASO #5: Se guarda el archivo: " + fileName);
                                Log("#4| Se guarda el scrap de Nro Documento de la imagen " + (o + 1) + ": " + fileName);

                                if (IsDataImage(fileName))
                                {
                                    Methods.LogProceso("#4: Validacion de imagen OK: " + fileName);
                                    approved = true;

                                    if (!FoundHeader)
                                    {
                                        lstBECTSImagePages.Add(new BECTSImagePage { CodePage = files[o].Name, isHeader = true });
                                        Col2ImagesPagesTemp.Add(new BECTSImagePage
                                        {
                                            CodePage = files[o].Name,
                                            isHeader = true,
                                            ImagePagesScraps = new List<BECTSScrap>()
                                        {
                                            new BECTSScrap { ScrapPath = fileName }
                                        }
                                        });

                                        FoundHeader = true;
                                    }
                                    else
                                    {
                                        lstBECTSImagePages.Add(new BECTSImagePage { CodePage = files[o].Name });
                                        Col2ImagesPagesTemp.Add(new BECTSImagePage
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
                                    Log("#4| Validacion NOK se procede a rotar: " + files[o].Name);
                                    ImageRotate(image, files[o].FullName.Replace(@"\ORIGINAL", ""));
                                }

                                //Me aseguro que termine la imagen en horizontal.
                                /*if (u == 3)
                                {
                                    if (image.Width < image.Height)
                                    {
                                        ImageRotate(image, files[o].FullName.Replace(@"\ORIGINAL", ""));
                                    }

                                }*/
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

                        approved = false;
                    }
                    else
                    {
                        /*if (DisapprovedFormats.Find(x => x.FormatCode == int.Parse(directoryName)) != null)
                        {*/
                        DisapprovedFormats.Add(new BECTSFormat
                        {
                            FormatCode = int.Parse(directoryName),
                            LstBECTSImagePages = lstBECTSImagePages
                        });
                        /*}*/
                    }

                }
            }
            catch (Exception exc)
            {
                Log("#4| ERROR: " + exc);
            }
            Log("#4| Termina rutina #4");
        }

        private void GetScrapsDataImages()
        {
            Log("#5| Inicia rutina #5 Obtiene Scraps de formatos aprobados"); //de manera automática
            Log("#5| Formatos aprobados: " + ApprovedFormats.Count);

            foreach (BECTSFormat formats in ApprovedFormats)
            {
                Log("#5| Formato a procesar: " + formats.FormatCode);

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
                                //newImage = Methods.cropAtRect(image, new Rectangle(2200, 200, 1500, 700));
                                newImage = Methods.cropAtRect(image, new Rectangle(1100, 100, 750, 350));
                                fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_Header_RUC" + Path.GetExtension(imagePages.CodePage);
                                newImage.Save(fileName);
                                imagePages.ImagePagesScraps.Add(new BECTSScrap { Type = "RUC", isHeaderValue = true, ScrapPath = fileName });

                                //CODIGO_LISTADO
                                //newImage = Methods.cropAtRect(image, new Rectangle(7480, 200, image.Width - 7480, 700));
                                newImage = Methods.cropAtRect(image, new Rectangle(3740, 100, image.Width - 3740, 350));
                                fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_Header_FormatCode" + Path.GetExtension(imagePages.CodePage);
                                newImage.Save(fileName);
                                imagePages.ImagePagesScraps.Add(new BECTSScrap { Type = "CODIGO_LISTADO", isHeaderValue = true, ScrapPath = fileName });

                                //TODO:PUEDE OBTENERSE EL PERIODO DE OTRA MANERA.
                                //PERIODO
                                //newImage = Methods.cropAtRect(image, new Rectangle(2200, 340, 1500, 700));
                                newImage = Methods.cropAtRect(image, new Rectangle(1100, 170, 750, 350));
                                fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_Header_Period" + Path.GetExtension(imagePages.CodePage);
                                newImage.Save(fileName);
                                imagePages.ImagePagesScraps.Add(new BECTSScrap { Type = "PERIODO", isHeaderValue = true, ScrapPath = fileName });

                                //AÑO
                                //newImage = Methods.cropAtRect(image, new Rectangle(4880, 340, 1500, 700));
                                newImage = Methods.cropAtRect(image, new Rectangle(2440, 170, 750, 350));
                                fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_Header_Year" + Path.GetExtension(imagePages.CodePage);
                                newImage.Save(fileName);
                                imagePages.ImagePagesScraps.Add(new BECTSScrap { Type = "ANIO", isHeaderValue = true, ScrapPath = fileName });

                                //MONEDA ABONO
                                //newImage = Methods.cropAtRect(image, new Rectangle(7480, 200, image.Width - 7480, 860));
                                newImage = Methods.cropAtRect(image, new Rectangle(3740, 100, image.Width - 3740, 430));
                                fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_Header_Currency" + Path.GetExtension(imagePages.CodePage);
                                newImage.Save(fileName);
                                imagePages.ImagePagesScraps.Add(new BECTSScrap { Type = "MONEDA", isHeaderValue = true, ScrapPath = fileName });
                            }

                            //COLUMNA 2 - NOMBRES Y APELLIDOS
                            //newImage = Methods.cropAtRect(image, new Rectangle(1570, 1494, 350, image.Height - 1494)); 
                            if (imagePages.isHeader)
                            {
                                //newImage = Methods.cropAtRect(image, new Rectangle(900, 2988, 1700, image.Height - 2988));
                                newImage = Methods.cropAtRect(image, new Rectangle(450, 1494, 850, image.Height - 1494));
                            }
                            else
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(450, 20, 850, image.Height - 20));
                            }
                            fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_2" + Path.GetExtension(imagePages.CodePage);
                            newImage.Save(fileName);
                            imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = false, ScrapPath = fileName, Type = "NombresYApellidos" });

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
                            imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = false, ScrapPath = fileName, Type = "TipDoc" });

                            //COLUMNA 4 - NRO DOCUMENTO
                            imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = false, Type = "NroDoc", ScrapPath = Col2ImagesPagesTemp.Where(x => x.CodePage == imagePages.CodePage).Select(x => x.ImagePagesScraps[0].ScrapPath).First() });

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
                            imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = false, ScrapPath = fileName, Type = "NroCta" });

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
                            imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = false, ScrapPath = fileName, Type = "CtaMoneda" });

                            //COLUMNA 7 - ABONO > MONTO
                            if (imagePages.isHeader)
                            {
                                //newImage = Methods.cropAtRect(image, new Rectangle(5320, 2988, 780, image.Height - 2988));
                                newImage = Methods.cropAtRect(image, new Rectangle(2660, 1494, 390, image.Height - 1494));
                            }
                            else
                            {
                                newImage = Methods.cropAtRect(image, new Rectangle(2660, 10, 390, image.Height - 10));
                            }
                            fileName = ScrapsDirectory + formats.FormatCode + @"\" + Path.GetFileNameWithoutExtension(imagePages.CodePage) + "_7" + Path.GetExtension(imagePages.CodePage);
                            newImage.Save(fileName);
                            imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = false, ScrapPath = fileName, Type = "AbonoMonto" });

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
                            imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = false, ScrapPath = fileName, Type = "AbonoMoneda" });

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
                            imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = false, ScrapPath = fileName, Type = "RemuMonto" });

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
                            imagePages.ImagePagesScraps.Add(new BECTSScrap { isHeaderValue = false, ScrapPath = fileName, Type = "RemuMoneda" });

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
        //            List<string> dataCol = GetData(t);

        //            t = formats.LstBECTSImagePages[0].CodePage + "_0_Header_FormatCode";
        //            dataCol = GetData(t);

        //            t = formats.LstBECTSImagePages[0].CodePage + "_0_Header_Period";
        //            dataCol = GetData(t);
        //            t = formats.LstBECTSImagePages[0].CodePage + "_0_Header_Year";
        //            dataCol = GetData(t);
        //            t = formats.LstBECTSImagePages[0].CodePage + "_0_Header_Currency";
        //            dataCol = GetData(t);
        //        }
        //    }
        //}

        //TODO: falta extraer los que tienen la tabla del tamaño de toda la hoja

        /*METODO GENERICO QUE RETORNO EL VALOR DE LA CONSULTA DEL CAMPO DE LA TABLA ASIGNADA*/
        private static string GetValor(string valor, string campo)
        {
            //using IBM.Data.DB2.iSeries;
            string id2 = null;
            /*CONECTANDO  A IBS*/
            valor = "47071029";
            campo = "cusidn";
            Console.WriteLine("Conectando a AS400 DB2 usando  metodo GetValor");
            iDB2Connection conn = null;
            try
            {
                conn = new iDB2Connection("DataSource=172.19.18.25;UserID=BFPFELQUI;Password=BFPFELQUI5;DataCompression=True;");
                conn.Open();
                if (conn != null)
                {
                    Console.WriteLine("Successfully connected...");
                    string qry = "select distinct(" + campo + ")  from bfpcyfiles.acmst a, bfpcyfiles.cumst b where acmaty = 'AHSF' and acmast<>'C' and a.acmcun = b.cuscun and cusidn='" + valor + "'";
                    iDB2Command comm = conn.CreateCommand();
                    comm.CommandText = qry;
                    Console.WriteLine("Entrará a metodo iDB2DataReader...");
                    iDB2DataReader reader = comm.ExecuteReader();
                    Console.WriteLine("Entrará a leer...");
                    //int i = 0;
                    if (reader.Read())
                    {
                        //TODO: atentos por si devuelve más de uno, igual se devuelve índice 0.
                        id2 = reader[0].ToString();//devuelve un valor del campo respectivo
                    }
                    reader.Close();
                    comm.Dispose();
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : " + ex);
                Console.WriteLine(ex.StackTrace);
            }
            finally
            {
                conn.Close();
            }
            if (null != id2)
            {
                return valor;
            }
            else
            {
                return "";
            }
        }

        private void getValue(SpreadsheetDocument document, List<String> dni, String col, String field)
        {
            foreach (string value in dni)
            {
                uint index = 25;
                //llamar funcion dask
                String val = GetValor(value, field);
                UpdateCell(document, val, index, col);
                index++;
            }
        }
        //TODO mandar a util. / ver si al final va a servir o no.
        //public static IEnumerable<T> SortAndDedupe<T>(this IEnumerable<T> input)
        //{
        //    var toDedupe = input.OrderBy(x => x);

        //    T prev = default(T);
        //    foreach (var element in toDedupe)
        //    {
        //        //if (element == prev) continue;
        //        if (EqualityComparer<T>.Default.Equals(element, prev)) continue;

        //        yield return element;
        //        prev = element;
        //    }
        //}

        private void GetDataManualFromScraps()
        {
            Log("#6.1| Inicia rutina #6 Obtiene data de los Scraps");

            DirectoryInfo diManual = new DirectoryInfo(ManualScrapsDirectory);
            DirectoryInfo[] directories = diManual.GetDirectories("*_OK");
            try
            {
                for (int i = 0; i < directories.Length; i++)
                {
                    /*foreach (BECTSFormat formats in ApprovedFormats)
                    {*/

                    string formatCode = directories[i].Name.Replace("_OK", "");

                    //COPIO PLANTILLA CON NOMBRE DE FORMATO.
                    string filename = directories[i].FullName + @"\PLANTILLA\" + "Plantilla abono CTS" + " - " + formatCode + ".xlsx";

                    //TODO: validar si el archivo existe y marcarlo en el log.
                    File.Copy(TemplatesDirectory + "Plantilla abono CTS.xlsx", filename, true);

                    directories[i].MoveTo(directories[i].FullName.Replace("_OK", "_PROCESANDO"));
                    //AL CAMBIAR A PROCESANDO, ACTUALIZAR FILENAME
                    filename = directories[i].FullName + @"\PLANTILLA\" + "Plantilla abono CTS" + " - " + formatCode + ".xlsx";


                    uint actualRow = 0;
                    /*uint nextTemp = 0;
                    uint next = 0;*/

                    List<BECTSFormatDetail> lst = new List<BECTSFormatDetail>();
                    List<string> dataCol = null;
                    FileInfo[] files = directories[i].GetFiles("NroDoc*.png");
                    for (int o = 0; o < files.Length; o++)
                    {
                        //actualRow += next;
                        dataCol = GetData(files[o].FullName);
                        dataCol = dataCol.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                        SetScrapsDetailData(filename, ref actualRow, ref lst, dataCol, formatCode, "NroDoc");
                        actualRow += 2;
                    }
                    //nextTemp = actualRow;
                    actualRow = 0;

                    files = directories[i].GetFiles("NombresYApellidos*.png");
                    for (int o = 0; o < files.Length; o++)
                    {
                        //actualRow += next;
                        dataCol = GetData(files[o].FullName);
                        dataCol = dataCol.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                        SetScrapsDetailData(filename, ref actualRow, ref lst, dataCol, formatCode, "NombresYApellidos");
                        actualRow += 2;
                    }
                    actualRow = 0;

                    files = directories[i].GetFiles("TipDoc*.png");
                    for (int o = 0; o < files.Length; o++)
                    {
                        //actualRow += next;
                        dataCol = GetData(files[o].FullName);
                        dataCol = dataCol.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                        SetScrapsDetailData(filename, ref actualRow, ref lst, dataCol, formatCode, "TipDoc");
                        actualRow += 2;
                    }
                    actualRow = 0;

                    files = directories[i].GetFiles("NroCta*.png");
                    for (int o = 0; o < files.Length; o++)
                    {
                        dataCol = GetData(files[o].FullName);
                        dataCol = dataCol.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                        SetScrapsDetailData(filename, ref actualRow, ref lst, dataCol, formatCode, "NroCta");
                        actualRow += 2;
                    }
                    actualRow = 0;

                    files = directories[i].GetFiles("CtaMoneda*.png");
                    for (int o = 0; o < files.Length; o++)
                    {
                        dataCol = GetData(files[o].FullName);
                        dataCol = dataCol.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                        SetScrapsDetailData(filename, ref actualRow, ref lst, dataCol, formatCode, "CtaMoneda");
                        actualRow += 2;
                    }
                    actualRow = 0;

                    files = directories[i].GetFiles("AbonoMonto*.png");
                    for (int o = 0; o < files.Length; o++)
                    {
                        dataCol = GetData(files[o].FullName);
                        dataCol = dataCol.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                        SetScrapsDetailData(filename, ref actualRow, ref lst, dataCol, formatCode, "AbonoMonto");
                        actualRow += 2;
                    }
                    actualRow = 0;

                    files = directories[i].GetFiles("AbonoMoneda*.png");
                    for (int o = 0; o < files.Length; o++)
                    {
                        dataCol = GetData(files[o].FullName);
                        dataCol = dataCol.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                        SetScrapsDetailData(filename, ref actualRow, ref lst, dataCol, formatCode, "AbonoMoneda");
                        actualRow += 2;
                    }
                    actualRow = 0;

                    files = directories[i].GetFiles("RemuMonto*.png");
                    for (int o = 0; o < files.Length; o++)
                    {
                        dataCol = GetData(files[o].FullName);
                        dataCol = dataCol.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                        SetScrapsDetailData(filename, ref actualRow, ref lst, dataCol, formatCode, "RemuMonto");
                        actualRow += 2;
                    }
                    actualRow = 0;

                    files = directories[i].GetFiles("RemuMoneda*.png");
                    for (int o = 0; o < files.Length; o++)
                    {
                        dataCol = GetData(files[o].FullName);
                        dataCol = dataCol.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                        SetScrapsDetailData(filename, ref actualRow, ref lst, dataCol, formatCode, "RemuMoneda");
                        actualRow += 2;
                    }
                    actualRow = 0;
                    //next = nextTemp;
                    directories[i].MoveTo(directories[i].FullName.Replace("_PROCESANDO", "_MANUAL"));

                    DisapprovedFormats.Find(x => x.FormatCode.ToString().Equals(formatCode)).Done = true;
                }
            }
            catch (Exception exc)
            {
                Log("#6.1| ERROR: " + exc.ToString());
            }
            Log("#6.1| Termina rutina #6.1 ");
        }

        private void GetDataFromScraps()
        {
            Log("#6| Inicia rutina #6 Obtiene data de los Scraps");
            List<String> dniList = new List<String>();

            foreach (BECTSFormat formats in ApprovedFormats)
            {
                //COPIO PLANTILLA CON NOMBRE DE FORMATO.
                string filename = ScrapsDirectory + formats.FormatCode + @"\PLANTILLA\" + "Plantilla abono CTS" + " - " + formats.FormatCode + ".xlsx";

                //TODO: validar si el archivo existe y marcarlo en el log.
                File.Copy(TemplatesDirectory + "Plantilla abono CTS.xlsx", filename, true);

                uint actualRow = 0;
                uint lastRowCountImagePage = 0;
                uint lastRowCountImagePageTemp = 0;

                // getValue(document, dni, col, "tipodoc");

                foreach (BECTSImagePage imagePages in formats.LstBECTSImagePages)
                {
                    //int initialCol = 2;
                    //string col = "";
                    //TODO: el salto de lineas en blanco debe ser parámetro, o hago método que evalue q tanto se parece la cadena al nombre que devuelve la búsqueda por dni en la reniec. o ignoro la columna nombres y apellidos desde el pdf y la pinto desde reniec.
                    //le agregamos más 2  espacios en blanco en caso se salte uno o máximo dos por error.
                    lastRowCountImagePage += lastRowCountImagePageTemp;

                    foreach (BECTSScrap scrap in imagePages.ImagePagesScraps)
                    {
                        List<string> dataCol = GetData(scrap.ScrapPath);

                        if (!scrap.isHeaderValue)
                        {
                            actualRow = 0;

                            //dataCol = dataCol.Where(s => !string.IsNullOrWhiteSpace(s)).SortAndDedupe().ToList();
                            //dtList = dtList.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList()
                            dataCol = dataCol.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();

                            List<BECTSFormatDetail> lst = formats.LstBECTSFormatDetail;
                            SetScrapsDetailData(filename, ref actualRow, ref lst, dataCol, formats.FormatCode.ToString(), scrap.Type);

                            formats.LstBECTSFormatDetail = lst;
                        }
                        else
                        {
                            //List<string> dataCol = GetData(scrap.ScrapPath);
                            SetScrapsHeaderData(filename, dataCol, formats.FormatCode.ToString(), scrap.Type);
                        }
                    }
                }

                formats.Done = true;
            }
            Log("#6| Termina rutina #6 ");
        }

        private void SetScrapsHeaderData(string fileName, List<string> dataCol, string formatCode, string scrapType)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
            {
                try
                {
                    switch (scrapType)
                    {
                        case "RUC":
                            UpdateCell(document, repairRUC(dataCol), 5, "D"); break;
                        case "CODIGO_LISTADO":
                            UpdateCell(document, formatCode, 5, "K"); break;
                        case "PERIODO":
                            UpdateCell(document, repairPeriod(dataCol), 7, "D"); break;
                        case "ANIO":
                            UpdateCell(document, repairYear(dataCol), 7, "H"); break;
                        case "MONEDA":
                            UpdateCell(document, repairCurrency(dataCol), 7, "K"); break;
                    }
                }
                catch (Exception exc)
                {
                    Log("#6| SetScrapsHeaderData > - ERROR: " + exc.ToString());
                }
            }
        }

        private void SetScrapsDetailData(string fileName, ref uint rowIndex, ref List<BECTSFormatDetail> formatDetails, List<string> dataCol, string formatCode, string scrapType)
        {
            string col = "";

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
            {
                try
                {


                    //completar los campos tipo documento,numero de cuenta, moneda cuenta, nombre
                    //getValue(document, dniList, "C", "cusna1"); // nombres y apellidos
                    //getValue(document, dniList, "D", "custid"); // tipo documento
                    //getValue(document, dniList, "F", "acmacc"); // cuenta
                    //getValue(document, dniList, "G", "acmccy"); // moneda cuenta

                    //TODO: estamos suponiendo que la cantidad de filas no difiere entre columnas
                    for (int i = 0; i < dataCol.Count; i++)
                    {
                        if (formatDetails.ElementAtOrDefault(i) == null)
                        {
                            formatDetails.Add(new BECTSFormatDetail());
                        }

                        string dataRepaired = "";

                        switch (scrapType)
                        {
                            case "NombresYApellidos":
                                col = "C";
                                dataRepaired = dataCol[i];
                                break;
                            case "TipDoc":
                                col = "D";
                                dataRepaired = dataCol[i];
                                break;
                            case "NroDoc":
                                col = "E";
                                dataRepaired = repairID(formatDetails[i].NombresYApellidos, dataCol[i]);
                                /*if ("" != dataRepaired)
                                {
                                    //isValid = true;
                                    //dniList.Add(dataRepaired); //se guardan los dni
                                }*/
                                break;
                            case "NroCta":
                                col = "F";
                                dataRepaired = repairAccountNumber("", dataCol[i]);
                                break;
                            case "CtaMoneda":
                                col = "G";
                                dataRepaired = dataCol[i];
                                break;
                            case "AbonoMonto": //CODIGO LISTADO
                                col = "H";
                                dataRepaired = repairAmount(dataCol[i]);
                                break;
                            case "AbonoMoneda": //PERIODO
                                col = "I";
                                dataRepaired = dataCol[i];
                                break;
                            case "RemuMonto": //PERIODO
                                col = "J";
                                dataRepaired = repairAmount(dataCol[i]);
                                break;
                            case "RemuMoneda": //PERIODO
                                col = "K";
                                dataRepaired = dataCol[i];
                                break;
                        }

                        var propertyInfo = formatDetails[i].GetType().GetProperty(scrapType);
                        try
                        {
                            dataRepaired = Convert.ChangeType(dataRepaired, propertyInfo.PropertyType).ToString();
                            propertyInfo.SetValue(formatDetails[i], Convert.ChangeType(dataRepaired, propertyInfo.PropertyType), null);
                        }
                        catch
                        {
                            dataRepaired = "0.0";
                            propertyInfo.SetValue(formatDetails[i], Convert.ChangeType(dataRepaired, propertyInfo.PropertyType), null);
                        }
                        UpdateCell(document, dataRepaired, 25 + rowIndex, col);
                        rowIndex++;

                        /*
                        initialCol++;

                        if (col == "E")
                        {
                            lastRowCountImagePageTemp = actualRow + 2;
                        }*/

                        //formatDetails[i]."" = "";
                    }
                }
                catch (Exception exc)
                {
                    Log("#6| > SetScrapsDetailData - ERROR: " + exc.ToString());
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

        public static int Compute(string s, string t)
        {
            int n = s.Length;
            int m = t.Length;
            int[,] d = new int[n + 1, m + 1];

            // Step 1
            if (n == 0)
            {
                return m;
            }

            if (m == 0)
            {
                return n;
            }

            // Step 2
            for (int i = 0; i <= n; d[i, 0] = i++)
            {
            }

            for (int j = 0; j <= m; d[0, j] = j++)
            {
            }

            // Step 3
            for (int i = 1; i <= n; i++)
            {
                //Step 4
                for (int j = 1; j <= m; j++)
                {
                    // Step 5
                    int cost = (t[j - 1] == s[i - 1]) ? 0 : 1;

                    // Step 6
                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                        d[i - 1, j - 1] + cost);
                }
            }
            // Step 7
            return d[n, m];
        }

        private string repairID(string names, string ID)
        {
            string newID = Regex.Replace(ID, @"[^\d]", String.Empty).PadLeft(8, '0');

            if (Regex.IsMatch(newID, @"^\d+$"))
            {
                if (Regex.IsMatch(newID, @"\d{8}"))
                {
                    //TODO, si machea regex igual evaluar si existe dni con reniec... 


                    //getValue(document, dniList, "C", "cusna1"); // nombres y apellidos

                    //GetValor(newID, "cusna1");
                    //int i = Compute(names, GetValor(newID, "cusna1"));

                    return newID;
                }
            }
            else
            {
                return "";
            }
            //TODO, poner rojo si no cumple.


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
                            Log("#6| " + page.GetMeanConfidence() * 100 + "% de Fiabilidad de imagen: " + Path.GetFileName(scrapPath));
                            //TODO: DESDE ACÁ PODRÍAMOS ELIMINAR LOS BLANCOS... .
                            List<string> prev = text.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.RemoveEmptyEntries).ToList<string>();
                            return prev;
                        }
                    }
                }
            }
            catch (Exception e)
            {

            }

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
            lstIDs = lstIDs.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
            if (lstIDs.Count > 0)
            {
                //TODO: meter el minimunAccepted en appconfig
                Int16 minimunAccepted = (Int16)(lstIDs.Count / 2);
                Int16 accepted = 0;
                foreach (string ID in lstIDs)
                {
                    if (Regex.IsMatch(ID, @"\d{8}"))
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
