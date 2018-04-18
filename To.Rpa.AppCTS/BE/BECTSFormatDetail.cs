using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace To.Rpa.AppCTS.BE
{
    public class BECTSFormatDetail
    {
        public string Id { get; set; }
        public string NombresYApellidos { get; set; }
        public string TipDoc { get; set; }
        public UInt16 NroDoc { get; set; }

        public UInt16 NroCta { get; set; }
        public string CtaMoneda { get; set; }
        public decimal AbonoMonto { get; set; }
        public decimal AbonoMoneda { get; set; }
        public decimal RemuMonto { get; set; }
        public string RemuMoneda { get; set; }

    }
}
