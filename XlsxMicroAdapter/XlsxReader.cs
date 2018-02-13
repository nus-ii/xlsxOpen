using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace XlsxMicroAdapter
{
    public class XlsxReader : IDisposable
    {
        //public Stream targetStream { get; set; }

        public MicroWorkbook Book { get; set; }

        
        //public XlsxReader(FileStream sourceStream)
        //{
        //    this.Book = new MicroWorkbook(sourceStream);
        //}

        public XlsxReader(string path)
        {
            this.Book = new MicroWorkbook(path);
        }

        public void AddList(string name)
        {
            MicroSheet newList = new MicroSheet(name);
            this.Book.Sheets.Add(newList);
        }

        public void Dispose()
        {
           // Book.Dispose();
            // targetStream.Dispose();
        }

        
    }

}//end of name space

