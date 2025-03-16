using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BillOfLading
{
    public class BLData
    {
        public int Docnum { get; set; }
    }
    public class PoData
    {
        public int DocEntry { get; set; }
        public int LineNum { get; set; }
        public double Qty { get; set; }
        public string ItemCode { get; set; }
        public double Avaiable { get; set; }
        public double BLAlloctaed { get; set; }
    }
    public class BillofLadingHeader
    {
        public string Vendor { get; set; }
        public string BLNumber { get; set; }
        public List<BillofLadingLine> lines { get; set; }
    }

    public class BillofLadingLine
    {
        public int DocEntry { get; set; }
        public int BaseEntry { get; set; }
        public int BaseLine { get; set; }
        public double Qty { get; set; }
        public string ItemCode { get; set; }
        public double UnitPrice { get; set; }
    }

    //List for Copy data from model form
    public class CopyData
    {
        public string ItemCode { get; set; }
        public double Qty { get; set; }
        public string BLNo { get; set; }
       public List<string> containerNo { get; set; }
    }
}
