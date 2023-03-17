using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Org.BouncyCastle.Bcpg.OpenPgp;

namespace ExcelAPP
{
    class SmetaFile : IComparable<SmetaFile>
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public string NameDate { get; set; }
        public string Price { get; set; }
        public int PageCount { get; set; }
        public FileInfo FolderInfo { get; set; }
        public string ShortCode { get; set; }
        public int Part { get; set; }
        public int NumOfPage { get; set; }

        public SmetaFile(string Code, string Name, string NameDate, string Price, int PageCount, FileInfo FolderInfo, string ShortCode)
        {
            this.Code = Code;
            this.Name = Name;
            this.NameDate = NameDate;
            this.Price = Price;
            this.PageCount = PageCount;
            this.FolderInfo = FolderInfo;
            this.ShortCode = ShortCode;
            this.Part = -1;
            this.NumOfPage = -1;
        }

        //public ~SmetaFile() //TODO make a destructor (garbage collection)
        //{
        //    this.Code = null;
        //    this.Name = null;
        //    this.Price = null;
        //    //this.PageCount = null; //TODO reset don't work 

        //    //GC.Collect();
        //    //GC.WaitForPendingFinalizers();
        //    //GC.Collect();
        //}



        public int CompareTo(SmetaFile other)
        {
            return other.Code.CompareTo(this.Code);

        }

        public override string ToString()
        {
            return $"Code: {Code} \nName: {Name} \nPrice: {Price} \nPageCount: {PageCount} \nFolderInfo{FolderInfo} \nShortCode: {ShortCode} \n\n";
            //return $"{Code} - {Name} - {FolderInfo}";
        }
    }



    public class Pair
    {
        public string Key { get; set; }
        public string Value { get; set; }
    }
}