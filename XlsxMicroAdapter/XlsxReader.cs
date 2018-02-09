using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;


namespace XlsxMicroAdapter
{
	public class XlsxReader:IDisposable
	{
		//public Stream targetStream { get; set; }

		public MicroWorkbook Book { get; set; }

		public XlsxReader()
		{
			this.Book=new MicroWorkbook();
		}

		public XlsxReader(Stream sourceStream)
		{
			//this.targetStream = sourceStream;
			this.Book = new MicroWorkbook(sourceStream);
		}

		public void AddList(string name)
		{
			MicroSheet newList=new MicroSheet(name);
			this.Book.Sheets.Add(newList);
		}

        public void Dispose()
        {
            Book.Dispose();
           // targetStream.Dispose();
        }

        public static int l(Stream sourceStream)
        {
            var Book = new MicroWorkbook(sourceStream);
            return Book.Sheets.Count;
        }
    }

}//end of name space

