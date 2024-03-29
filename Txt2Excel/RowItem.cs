﻿using System;

namespace Txt2Excel
{
    internal class RowWithDate : IEquatable<RowWithDate>, IComparable<RowWithDate>
    {
        public int Id { get; set; }
        public String BarCode { get; set; }
        public String Description { get; set; }
        public String Expiration { get; set; }
        public DateTime FormattedExpiration { get; set; }
        public String Month { get; set; }
        public String Year { get; set; }

        //Not parsable datas
        public RowWithDate(int Id, string BarCode, string Description, string Expiration, string Month, String Year)
        {
            this.Id = Id;
            this.BarCode = BarCode;
            this.Description = Description;
            this.Expiration = Expiration;
            this.Month = Month;
            this.Year = Year;
        }

        public RowWithDate(int Id, string BarCode, string Description, string Expiration, DateTime FormattedExpiration, string Month, String Year)
        {
            this.Id = Id;
            this.BarCode = BarCode;
            this.Description = Description;
            this.Expiration = Expiration;
            this.FormattedExpiration = FormattedExpiration;
            this.Month = Month;
            this.Year = Year;
        }


        public bool Equals(RowWithDate OtherRowItem)
        {
            if (this.BarCode.Equals(OtherRowItem.BarCode) &&
                   this.Description.Equals(OtherRowItem.Description) &&
                   this.Expiration.Equals(OtherRowItem.Expiration) &&
                   this.FormattedExpiration.Equals(OtherRowItem.FormattedExpiration))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public override bool Equals(Object Other)
        {
            if (Other == null)
            {
                return false;
            }
            RowWithDate OtherRowItem = Other as RowWithDate;
            if (OtherRowItem == null)
            {
                return false;
            }
            else
            {
                return this.Equals(OtherRowItem);
            }
        }

        //Compare by FormattedExpiration Asc
        public int CompareTo(RowWithDate OtherRowItem)
        {
            // A null value means that this object is greater.
            if (OtherRowItem == null)
                return 1;

            else
                return this.FormattedExpiration.CompareTo(OtherRowItem.FormattedExpiration);
        }

        public override int GetHashCode()
        {
            return Id;
        }
    }
}
