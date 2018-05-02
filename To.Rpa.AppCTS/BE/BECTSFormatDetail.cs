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
        public string NroDoc { get; set; }

        public string NroCta { get; set; }
        public string CtaMoneda { get; set; }
        public decimal AbonoMonto { get; set; }
        public string AbonoMoneda { get; set; }
        public decimal RemuMonto { get; set; }
        public string RemuMoneda { get; set; }

    }
}
