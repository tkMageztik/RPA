using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tesseract;

namespace To.AtNinjas.Karen
{
    public class KarenCore
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="scrapPath"></param>
        /// <returns></returns>
        public List<List<string>> GetData(string scrapPath)
        {
            try
            {
                using (var engine = new TesseractEngine(@"./Dictionaries/Tesseract/esp", "spa", EngineMode.Default))
                {

                    using (var img = Pix.LoadFromFile(scrapPath))
                    {
                        using (var page = engine.Process(img))
                        {
                            //TODO: parece que el primer registro lo lee como salto de línea.
                            var text = page.GetText();
                            //Log("#6| " + page.GetMeanConfidence() * 100 + "% de Fiabilidad de imagen: " + Path.GetFileName(scrapPath));
                            
                            List<string> prev = text.Split(new string[] { "\n", "\r\n" }, StringSplitOptions.None).ToList<string>();

                            List<List<string>> lst = new List<List<string>>();
                            List<string> lstChild = new List<string>();

                            foreach (string str in prev)
                            {
                                if (str.Equals(""))
                                {
                                    lstChild = lstChild.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                                    lstChild = lstChild.Where(s => !string.IsNullOrEmpty(s)).ToList();
                                    if (lstChild.Count != 0)
                                    {
                                        lst.Add(lstChild);
                                        lstChild = new List<string>();
                                    }
                                }
                                else
                                {
                                    lstChild.Add(str);
                                }
                            }
                            
                            return lst;
                        }
                    }
                }
            }
            catch (Exception e)
            {

            }

            return null;
        }
    }
}
