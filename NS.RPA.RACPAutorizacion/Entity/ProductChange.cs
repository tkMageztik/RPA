using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NS.RPA.RACPAutorizacion
{
    public class ProductChange
    {
        public string Plot { get; set; }
        public string FundingNumber { set; get; }
        public string CompanySize { set; get; }
        public string Product { set; get; }
        public string ChangeProduct { set; get; }
        public string State { set; get; }
        public override string ToString()
        {
            return Plot;
        }
    }
}
