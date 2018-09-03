using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Engine.Models
{
    class SO
    {
        public string Product { get; set; }
        public string JobNum { get; set; }
        public string PO { get; set; }
        public double Cost { get; set; }
        public double AddedValue { get; set; }
        public string UM { get; set; }
        public double Weight { get; set; }
        public string Factura { get; set; }

        public SO(string product, string jobNum, string po, double cost, double addedValue, double weight, string um, string factura)
        {
            Product = product;
            JobNum = jobNum;
            PO = po;
            Cost = cost;
            AddedValue = addedValue;
            Weight = weight;
            UM = um;
            Factura = factura;
        }
    }
}
