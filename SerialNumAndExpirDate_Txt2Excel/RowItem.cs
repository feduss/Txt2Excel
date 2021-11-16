using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BarCodeDescrExpirDate_Txt2Excel
{
    internal class RowItem
    {
        public String BarCode { get; set; }
        public String Description { get; set; }
        public String Expiration { get; set; }

        public RowItem(string BarCode, string Description, string Expiration)
        {
            this.BarCode = BarCode;
            this.Description = Description;
            this.Expiration = Expiration;
        }
    }
}
