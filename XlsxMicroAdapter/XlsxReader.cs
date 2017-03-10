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
	public class XlsxReader
	{
		public Stream targetStream { get; set; }

		public MicroWorkbook Book { get; set; }

		public XlsxReader()
		{
			this.Book=new MicroWorkbook();
		}

		public XlsxReader(Stream sourceStream)
		{
			this.targetStream = sourceStream;
			this.Book = new MicroWorkbook(this.targetStream);
		}

		public void AddList(string name)
		{
			MicroSheet newList=new MicroSheet(name);
			this.Book.Sheets.Add(newList);
		}

	}

}//end of name space

