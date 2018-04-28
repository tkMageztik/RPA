using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace To.Rpa.AppCTS.BE
{
    public class BECTSFormat
    {
        public int FormatCode { get; set; }
        public int Pages { get; set; }
        public List<BECTSImagePage> LstBECTSImagePages { get; set; }
        public List<BECTSFormatDetail> LstBECTSFormatDetail { get; set; }

        public BECTSFormat()
        {
            LstBECTSImagePages = new List<BECTSImagePage>();
            LstBECTSFormatDetail = new List<BECTSFormatDetail>();
        }
    }

    public class BECTSImagePage
    {
        public string CodePage { get; set; }
        //public List<string> ImagePagesScraps { get; set; }
        public List<BECTSScrap> ImagePagesScraps { get; set; }
        public bool isHeader { get; set; } = false;

        public BECTSImagePage()
        {
            ImagePagesScraps = new List<BECTSScrap>();

            //key identifica si es un headerValue, si es un campo del header (ruc,codigoListado, etc)
        }
    }

    public class BECTSScrap
    {
        public string ScrapPath { get; set; }
        public string Type { get; set; }
        public bool isHeaderValue { get; set; } = false;
    }
}
