using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ExcelApp
{
    class SmetaFile : IComparable<SmetaFile>
    {
        public string Code { get; set; }
        public string Name { get; set; }
        public string Price { get; set; }
        public int PageCount { get; set; }
        public FileInfo FolderInfo { get; set; }

        public SmetaFile(string Code, string Name, string Price, int PageCount, FileInfo FolderInfo)
        {
            this.Code = Code;
            this.Name = Name;
            this.Price = Price;
            this.PageCount = PageCount;
            this.FolderInfo = FolderInfo;

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
            //return $"{Code} - {Name} - {Price} - {PageCount}";
            return $"{Code} - {Name} - {FolderInfo}";
        }
    }
}
