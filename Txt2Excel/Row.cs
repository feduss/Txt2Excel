using System;

namespace Txt2Excel
{
    internal class Row
    {
        public int Id { get; set; }
        public String[] Values { get; set; }
        //Not parsable datas
        public Row(int Id, string[] Values)
        {
            this.Id = Id;
            this.Values = Values;
        }
    }
}
