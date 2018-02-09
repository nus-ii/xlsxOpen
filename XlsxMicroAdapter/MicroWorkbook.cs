using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Excel = DocumentFormat.OpenXml.Office.Excel;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace XlsxMicroAdapter
{
    public class MicroWorkbook : IDisposable
    {
        public string Name { get; set; }

        public List<MicroSheet> Sheets;


        public MicroWorkbook(string name = "")
        {
            Name = name;
            Sheets = new List<MicroSheet>();
        }

        public MicroWorkbook(Stream targetStream, string name = "") : this(name)
        {
            this.Sheets = ReadSheets(targetStream);
        }

        private static List<MicroSheet> ReadSheets(Stream targetStream)
        {
            var sheets = new List<MicroSheet>();
            try
            {

                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(targetStream, true))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    var u = workbookPart.GetPartsOfType<SharedStringTablePart>();
                    if (u != null && u.Count() != 0)
                    {
                        SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                        SharedStringTable sst = sstpart.SharedStringTable;
                        GetValue(workbookPart, sst, sheets, true);
                    }
                    else
                    {
                        SharedStringTable sst = new SharedStringTable();
                        GetValue(workbookPart, sst, sheets, false);
                    }
                }
                return sheets;
            }
            catch (Exception e)
            {
                targetStream.Close();
                throw e;
            }
        }

        private static void GetValue(WorkbookPart workbookPart, SharedStringTable sst, List<MicroSheet> sheets, bool useShared)
        {
            List<Tuple<string, string, bool>> TupleList = new List<Tuple<string, string, bool>>();

            var lists = workbookPart.Workbook.Descendants<Sheet>();
            foreach (var l in lists)
            {
                bool visible = true;
                if (l.State != null)
                    visible = SheetStateValues.Visible == l.State.Value;

                var tempTuple = new Tuple<string, string, bool>(l.Id, l.Name, visible);
                TupleList.Add(tempTuple);
            }

            foreach (var idNamePair in TupleList)
            {
                var worksheetPart = workbookPart.GetPartById(idNamePair.Item1) as WorksheetPart;

                var activSheet = new MicroSheet(idNamePair.Item2, idNamePair.Item3);
                var sheet = worksheetPart.Worksheet;
                var cells = sheet.Descendants<Cell>();

                foreach (var cell in cells)
                {
                    var activeCell = CellRead(cell, sst, useShared);
                    activSheet.AddCell(activeCell);
                }
                sheets.Add(activSheet);
            }
        }

        private static MicroCell CellRead(Cell cell, SharedStringTable sst, bool useShared)
        {
            MicroCell result = new MicroCell();
            result.Row = cell.CellReference.Value.GetInt().ToString();
            result.Column = cell.CellReference.Value.GetLetter();

            if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
            {

                if (cell.CellFormula == null)
                {
                    int ssid = int.Parse(cell.CellValue.Text);
                    result.FormulaValue = cell.CellFormula?.Text;
                    string str = sst.ChildElements[ssid]?.InnerText ?? "";
                    //TODO: Болше проверок на null

                    result.ViewValue = str; //str;
                }
                else
                {
                    result.FormulaValue = cell.CellFormula?.Text;
                }

            }
            else
            {
                if (string.IsNullOrEmpty(cell.CellFormula?.Text))
                {
                    if (!string.IsNullOrEmpty(cell.CellValue?.Text))
                    { result.ViewValue = cell.CellValue.Text; }
                    else
                    {
                        result.ViewValue = cell.InnerText;
                    }
                }
                else
                { result.FormulaValue = cell.CellFormula?.Text; }
            }
            return result;
        }

        public Stream WriteSheets()
        {
            try
            {
                Stream targetStream = new MemoryStream();
                using (SpreadsheetDocument myDoc = SpreadsheetDocument.Create(targetStream, SpreadsheetDocumentType.Workbook))
                {
                    CreateParts(myDoc, this.Sheets);
                }
                //SetActiveSheetByMem(ghjStream);
                return targetStream;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private void SetActiveSheetByMem(Stream mem)
        {
            using (SpreadsheetDocument rptTemplate = SpreadsheetDocument.Open(mem, true))
            {
                foreach (OpenXmlElement oxe in (rptTemplate.WorkbookPart.Workbook.Sheets).ChildElements)
                {
                    if (((DocumentFormat.OpenXml.Spreadsheet.Sheet)(oxe)).Name != "")
                    {
                        //((DocumentFormat.OpenXml.Spreadsheet.Sheet)(oxe)).State = SheetStateValues.Hidden;

                        WorkbookView wv = null;

                        if (rptTemplate.WorkbookPart.Workbook.BookViews != null)
                        { wv = rptTemplate.WorkbookPart.Workbook.BookViews.ChildElements.First<WorkbookView>(); }
                        else
                        {
                            var bv = new BookViews();

                        }

                        if (wv != null)
                        {
                            wv.ActiveTab = GetIndexOfFirstVisibleSheet(rptTemplate.WorkbookPart.Workbook.Sheets);
                        }
                    }
                }
                rptTemplate.WorkbookPart.Workbook.Save();
            }
        }

        private void CreateParts(SpreadsheetDocument myDoc, List<MicroSheet> sheets)
        {
            WorkbookPart workbookPart = myDoc.AddWorkbookPart();
            GenerateWorkbookPartContent(workbookPart, sheets);

            int i = 1;
            foreach (var s in sheets)
            {
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>(string.Format("rId{0}", i));
                GenerateWorksheetPartContent(worksheetPart, s);
                i++;
            }
            //SetActiveSheet(workbookPart1);
        }

        private void SetActiveSheet(WorkbookPart workbookPart1)
        {
            WorkbookView wv = workbookPart1.Workbook.BookViews.ChildElements.First<WorkbookView>();

            if (wv != null)
            {
                wv.ActiveTab = GetIndexOfFirstVisibleSheet(workbookPart1.Workbook.Sheets);
            }
        }

        private UInt32Value GetIndexOfFirstVisibleSheet(Sheets sheets)
        {
            uint index = 0;
            foreach (Sheet currentSheet in sheets.Descendants<Sheet>())
            {
                if (currentSheet.State == null || currentSheet.State.Value == SheetStateValues.Visible)
                {
                    return index;
                }
                index++;
            }
            throw new Exception("No visible sheet found.");
        }

        private void SetDataCheckArray(Worksheet worksheet1, MicroSheet microSheet)
        {
            WorksheetExtensionList worksheetExtensionList = new WorksheetExtensionList();

            foreach (var checkDataItem in microSheet.CheckList)
            {
                SetDataCheck(worksheet1, checkDataItem, ref worksheetExtensionList);
            }
            worksheet1.Append(worksheetExtensionList);
        }

        private void SetDataCheck(Worksheet ws, DataCheckInfo info, ref WorksheetExtensionList worksheetExtensionList)
        {

            WorksheetExtension worksheetExtension = new WorksheetExtension() { Uri = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" };
            worksheetExtension.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");

            X14.DataValidations dataValidations = new X14.DataValidations() { Count = (UInt32Value)3U };
            dataValidations.AddNamespaceDeclaration("xm", "http://schemas.microsoft.com/office/excel/2006/main");

            //sites validation
            string checkFormula = GetCheckFormula(info);
            dataValidations.Append(new X14.DataValidation()
            {
                Type = DataValidationValues.List,
                AllowBlank = true,
                ShowInputMessage = true,
                ShowErrorMessage = true,
                DataValidationForumla1 = new X14.DataValidationForumla1() { Formula = new Excel.Formula(checkFormula) },
                ReferenceSequence = new Excel.ReferenceSequence(string.Concat(info.Target.Column, info.Target.Row))
            });

            worksheetExtension.Append(dataValidations);
            worksheetExtensionList.Append(worksheetExtension);
        }

        private string GetCheckFormula(DataCheckInfo info)
        {
            //"car_firm!$E$2:$E$61"
            return string.Format("{0}!${3}${4}:${1}${2}", info.SourceSheetName, info.SourceTopLeft.Column, info.SourceTopLeft.Row, info.SourceBottomRight.Column, info.SourceBottomRight.Row);
        }

        private void GenerateWorksheetPartContent(WorksheetPart worksheetPart, MicroSheet microSheet)
        {
            Worksheet worksheet = new Worksheet();
            SheetData sheetData = new SheetData();

            var xlsxRows = new Dictionary<int, Row>();
            var microRows = microSheet.RowsInt;
            var allCells = microSheet.GetAllCells();

            foreach (var microRow in microRows)
            {
                xlsxRows.Add(microRow, new Row());
            }

            foreach (var c in allCells)
            {
                var targetRow = xlsxRows[c.Value.RowInt];
                targetRow.Append(MicroCellToXlsxCell(c.Value));
            }

            foreach (var xlsxRow in xlsxRows)
            {
                sheetData.Append(xlsxRow.Value);
            }

            worksheet.Append(sheetData);
            SetDataCheckArray(worksheet, microSheet);
            worksheetPart.Worksheet = worksheet;
        }

        private static Cell MicroCellToXlsxCell(MicroCell microCell)
        {
            Cell cell1 = new Cell() { CellReference = string.Concat(microCell.Column, microCell.Row), DataType = CellValues.InlineString };
            InlineString inlineString1 = new InlineString();
            Text text1 = new Text();
            text1.Text = microCell.ViewValue;
            inlineString1.Append(text1);

            if (!string.IsNullOrEmpty(microCell.FormulaValue))
            {
                cell1.CellFormula = new CellFormula(microCell.FormulaValue);
            }
            else
            {
                cell1.Append(inlineString1);
            }
            return cell1;
        }

        private void GenerateWorkbookPartContent(WorkbookPart workbookPart1, List<MicroSheet> sheets)
        {
            Workbook workbook = new Workbook();
            workbook.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            Sheets sheets1 = new Sheets();

            uint i = 1;
            foreach (var s in sheets)
            {
                Sheet sheet1 = new Sheet() { Name = s.Name, SheetId = i, Id = string.Format("rId{0}", i) };

                if (!s.Visible)
                    sheet1.State = SheetStateValues.Hidden;

                sheets1.Append(sheet1);
                i++;
            }
            workbook.Append(sheets1);
            workbookPart1.Workbook = workbook;
        }

        public void Dispose()
        {
            foreach (var s in this.Sheets)
            {
                s.Dispose();
            }
            Sheets = null;
        }
    }//end of class

    internal static class StringHelper
    {
        public static string GetLetter(this string target)
        {
            string result = "";

            foreach (var l in target)
            {
                if (char.IsLetter(l))
                    result = string.Concat(result, l);
            }
            return result;
        }

        public static int GetInt(this string target)
        {
            string temp = "";
            foreach (var l in target)
            {
                if (char.IsDigit(l))
                    temp = string.Concat(temp, l);
            }
            return Convert.ToInt32(temp);
        }
    }
}
