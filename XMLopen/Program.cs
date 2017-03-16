﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using XlsxMicroAdapter;
//Получение данных их xlsx файла

namespace XMLopen
{
	class Program
	{
		static void Main(string[] args)
		{
			for (int i = 1; i < 650; i++)
			{
				Console.Write(String.Concat(i, "-", XlsxHelper.GetColumnLetter(i), " "));
			}

			//using (Stream fs = GetStream(@"C:\MyTempXls\cla2.xlsx"))
			//{
			//	var x=new XlsxReader(fs);
			//	var ghjkl=x.Book.WriteSheets();
				
			//	CopyStream(ghjkl, @"C:\MyTempXls\cla477.xlsx");
			//	ghjkl.Close();
			//}
			//using (Stream fs = GetStream(@"C:\MyTempXls\cla477.xlsx"))
			//{
			//	var x = new XlsxReader(fs);
			//	var ghjkl = x.Book.WriteSheets();
			//	CopyStream(ghjkl, @"C:\MyTempXls\cla477.xlsx");
			//	ghjkl.Close();
			//}
			MicroWorkbook mock = GetMockBook();
			var stream = mock.WriteSheets();
			string mockname = GetMockName();
			CopyStream(stream, mockname);
			Console.WriteLine(mockname);

			using (Stream fs = GetStream(mockname))
			{
				var x = new XlsxReader(fs);
				var ghjkl = x.Book.WriteSheets();

				CopyStream(ghjkl, @"C:\MyTempXls\cla477.xlsx");
				ghjkl.Close();
			}
			Console.ReadLine();
		}

		private static string GetMockName()
		{
			var n = DateTime.Now;
			string dt = string.Format("{0}-{1}-{2}", n.Hour, n.Minute, n.Second);
			return string.Concat(@"C:\MyTempXls\mock",dt,".xlsx");
		}

		private static MicroWorkbook GetMockBook()
		{
			var result=new MicroWorkbook();
			

			var aa = new MicroSheet("s");
			var qq = new MicroCell("1", "A", "testA");
			var ww = new MicroCell("2", "A", "testB");
			var ee = new MicroCell("3", "A", "testC");
			aa.AddCell(qq);
			aa.AddCell(ww);
			aa.AddCell(ee);
			result.Sheets.Add(aa);

			var a = new MicroSheet("14");
			var q = new MicroCell("1", "A", "testA");
			var w = new MicroCell("2", "A", "testA");
			var e = new MicroCell("3", "A", "testA");
			a.AddCell(q);
			a.AddCell(w);
			a.AddCell(e);

			a.CheckList.Add(new DataCheckInfo(qq,ee,"s",q));
			a.CheckList.Add(new DataCheckInfo(qq, ee, "s", w));
			a.CheckList.Add(new DataCheckInfo(qq, ee, "s", e));
			result.Sheets.Add(a);
			return result;
		}

		private static Stream GetStream(string path)
		{
			var result= new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);

			return (Stream) result;
		}

		public static void CopyStream(Stream inputStream, string destPath)
		{

			using (var fileStream = new FileStream(destPath, FileMode.Create, FileAccess.Write))
			{
				inputStream.Position = 0;
			    inputStream.CopyTo(fileStream);
				fileStream.Close();
			}
		}
	}
}