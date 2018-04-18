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
        public List<BECTSImagePages> LstBECTSImagePages { get; set; }
        public List<BECTSFormatDetail> LstBECTSFormatDetail { get; set; }

        public BECTSFormat()
        {
            LstBECTSImagePages = new List<BECTSImagePages>();
            LstBECTSFormatDetail = new List<BECTSFormatDetail>();
        }
    }

    public class BECTSImagePages
    {
        public string CodePage { get; set; }
        public List<string> ImagePagesScraps { get; set; }
        public bool isHeader { get; set; } = false;

        public BECTSImagePages()
        {
            ImagePagesScraps = new List<string>();
        }
    }
}
