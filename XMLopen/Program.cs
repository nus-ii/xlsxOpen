using System;
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

			using (Stream fs = GetStream(@"C:\MyTempXls\cla2.xlsx"))
			{
				var x=new XlsxReader(fs);
				var ghjkl=x.Book.WriteSheets();
				
				CopyStream(ghjkl, @"C:\MyTempXls\cla477.xlsx");
				ghjkl.Close();
			}
			using (Stream fs = GetStream(@"C:\MyTempXls\cla477.xlsx"))
			{
				var x = new XlsxReader(fs);
				//var ghjkl = x.Book.WriteSheets();
				//CopyStream(ghjkl, @"C:\MyTempXls\cla477.xlsx");
				//ghjkl.Close();
			}
			Console.WriteLine("All done!!!");
			Console.ReadLine();
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
