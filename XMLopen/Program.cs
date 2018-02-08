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
            #region MyRegion
            //MicroWorkbook mock = GetMockBook();
            //var stream = mock.WriteSheets();
            //string mockname = GetMockName();
            //CopyStream(stream, mockname);
            //Console.WriteLine(mockname);
            //Console.ReadLine();
            //using (Stream fs = GetStream(mockname))
            //{
            //    var x = new XlsxReader(fs);
            //    var ghjkl = x.Book.WriteSheets();

            //    CopyStream(ghjkl, @"C:\MyTempXls\momo.xlsx");
            //    ghjkl.Close();
            //}
            //Console.WriteLine("You can edit momo");
            //Console.ReadLine(); 
            #endregion


            //List<string> mockData = new List<string>();
            //mockData.Add("Part1;Part2;Part3;Part4;Part5;Part6;Part7;Part8;Part9;Part10;");

            //for (int i = 0; i < 1000; i++)
            //{
            //    var ga = Guid.NewGuid().ToString().Split('-');
            //    var gb = Guid.NewGuid().ToString().Split('-');

            //    mockData.Add(string.Format("{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};", ga[0], ga[1], ga[2], ga[3], ga[4], gb[0], gb[1], gb[2], gb[3], gb[4], gb[4]));

            //}

            //File.WriteAllLines(@"C:\MyTempXls\megamock1000.csv", mockData.ToArray());

            using (Stream fs = GetStream(@"C:\MyTempXls\megamock1000.xlsx"))
            {
                var x = new XlsxReader(fs);
                //var y = x.Book.Sheets.FirstOrDefault().HeadersDictionary;
                //var r = x.Book.Sheets.FirstOrDefault().GetCellsWhereRow("2");
                //var targetCell = x.Book.Sheets.FirstOrDefault().GetCellByHeader(2, "Part1").ViewValue;


                List<List<string>> readResult = new List<List<string>>();
                var tarS = x.Book.Sheets.FirstOrDefault();
                var tarR = tarS.RowsInt;
                var tarH = tarS.HeadersDictionary;

                foreach (var tr in tarR)
                {
                    List<string> tempL = new List<string>();
                    foreach (var th in tarH)
                    {
                        tempL.Add(tarS.GetCellByHeader(tr, th.Value).ViewValue);
                    }
                    readResult.Add(tempL);
                }

                //var columns = x.Book.Sheets.FirstOrDefault();
                //var tempReader = x.Book.WriteSheets();

                //CopyStream(tempReader, @"C:\MyTempXls\momo2.xlsx");
                //tempReader.Close();
                Console.WriteLine("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF");
            }

           

            Console.ReadLine();
        }

        private static string GetMockName()
        {
            var n = DateTime.Now;
            string dt = string.Format("{0}-{1}-{2}", n.Hour, n.Minute, n.Second);
            return string.Concat(@"C:\MyTempXls\mock", dt, ".xlsx");
        }

        private static MicroWorkbook GetMockBook()
        {
            var result = new MicroWorkbook();


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
            var r = new MicroCell("4", "A", "20", "=5*4");
            a.AddCell(q);
            a.AddCell(w);
            a.AddCell(e);
            a.AddCell(r);
            result.Sheets.Add(a);

            List<MicroCell> lis = new List<MicroCell>();
            for (int z = 5; z < 2000; z++)
            {
                var nnn = new MicroCell(z.ToString(), "B", "testA");
                lis.Add(nnn);
                a.CheckList.Add(new DataCheckInfo(qq, ee, "s", nnn));
            }
            a.AddCellList(lis);

            a.CheckList.Add(new DataCheckInfo(qq, ee, "s", q));
            a.CheckList.Add(new DataCheckInfo(qq, ee, "s", w));
            a.CheckList.Add(new DataCheckInfo(qq, ee, "s", e));

            return result;
        }

        private static Stream GetStream(string path)
        {
            var result = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);

            return (Stream)result;
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
