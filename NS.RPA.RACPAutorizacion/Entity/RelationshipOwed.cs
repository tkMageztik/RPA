using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NS.RPA.RACPAutorizacion
{
    //Relación de Adeudado
    public class RelationshipOwed
    {
        public string Plot { get; set; }
        public string ContractNumberDue { set; get; }
        public string FundingNumber { set; get; }
        public string State { set; get; }
        public override string ToString()
        {
            return Plot;
        }
    }
}
