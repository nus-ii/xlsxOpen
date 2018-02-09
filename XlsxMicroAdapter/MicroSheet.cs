using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace XlsxMicroAdapter
{
    public class MicroSheet:IDisposable
    {
        public string Name { get; set; }

        private bool Inited;

        public int HeaderRow
        {
            get
            {
                return HeaderRowValue;
            }
            set
            {
                HeaderRowValue = value;
                Init();                
            }
        }

        private int HeaderRowValue;

        private Dictionary<string, MicroCell> Cells;

        private Hashtable CellPacksSortedByRowHash;

        public bool Visible { get; set; }

        private List<int> RowsList;

        private List<string> ColumnsList;

        public List<DataCheckInfo> CheckList;

        public List<string> Rows
        {
            get
            {
                if (!Inited)
                    Init();

                var result = new List<string>();

                foreach (var r in RowsList)
                {
                    result.Add(r.ToString());
                }
                return result;
            }
        }

        public List<int> RowsInt
        {
            get
            {
                if (!Inited)
                    Init();

                return RowsList;
            }
        }

        public List<string> Columns
        {
            get
            {
                if (!Inited)
                    Init();

                return ColumnsList;
            }
        }

        public Dictionary<string, string> HeadersDictionary
        {
            get
            {
                if (!Inited)
                    Init();

                return HeadersDictionaryValue;
            }
        }

        private Dictionary<string, string> HeadersDictionaryValue;

        public MicroSheet(string name = "", bool visible = true)
        {
            this.Name = name;
            this.Cells = new Dictionary<string, MicroCell>();
            this.Visible = visible;
            this.ColumnsList = new List<string>();
            this.RowsList = new List<int>();
            this.CheckList = new List<DataCheckInfo>();
            this.HeaderRowValue = 1;
        }

        public void AddCell(MicroCell newCell)
        {
            this.Cells.Add(string.Concat(newCell.Column, newCell.Row), newCell);
            this.Inited = false;
        }

        private void Init()
        {
            FixColumnListAndRowList();
            HeadersDictionaryValue = GetHeadersDictionary();
            this.CellPacksSortedByRowHash = GetCellPacksSortedByRow();
            Inited = true;
        }

        private void FixColumnListAndRowList()
        {
            foreach (var cell in this.Cells)
            {
                if (!ColumnsList.Contains(cell.Value.Column))
                    ColumnsList.Add(cell.Value.Column);

                if (!RowsList.Contains(cell.Value.RowInt))
                    RowsList.Add(cell.Value.RowInt);
            }

        }

        private void FixRowList()
        {
            foreach (var cell in this.Cells)
            {
                if (!RowsList.Contains(cell.Value.RowInt))
                    RowsList.Add(cell.Value.RowInt);
            }
        }

        private void FixColumnList()
        {
            foreach (var cell in this.Cells)
            {
                if (!ColumnsList.Contains(cell.Value.Column))
                    ColumnsList.Add(cell.Value.Column);
            }
        }

        public void AddCellList(List<MicroCell> CellList)
        {
            foreach (var cell in CellList)
            {
                this.AddCell(cell);
            }
            Inited = false;
        }

        public List<MicroCell> GetCellsWhereRow(string row)
        {
            List<MicroCell> result = new List<MicroCell>();
            var preResult = this.Cells.Where(c => c.Value.Row == row).ToList();

            foreach (var cell in preResult)
            {
                result.Add(cell.Value);
            }
            return result;

        }

        public List<MicroCell> GetCellsWhereColumn(string column)
        {
            List<MicroCell> result = new List<MicroCell>();
            var preResult = this.Cells.Where(c => c.Value.Column == column).ToList();
            foreach (var cell in preResult)
            {
                result.Add(cell.Value);
            }
            return result;
        }

        public Hashtable GetCellPacksSortedByRow()
        {
            Hashtable result = new Hashtable(new Dictionary<int, List<MicroCell>>());
            foreach (var cell in this.Cells)
            {
                if (result.ContainsKey(cell.Value.RowInt))
                {
                    var t = (List<MicroCell>)result[cell.Value.RowInt];
                    t.Add(cell.Value);
                }
                else
                {
                    var nl = new List<MicroCell>();
                    nl.Add(cell.Value);
                    result.Add(cell.Value.RowInt, nl);
                }
            }
            return result;
        }


        /// <summary>
        /// Возвращает запрошенную ячейку
        /// </summary>
        /// <param name="row">строка</param>
        /// <param name="column">столбец</param>
        /// <returns></returns>
        public MicroCell GetCell(string row, string column)
        {
            string adress = string.Concat(column, row);
            MicroCell result = new MicroCell();
            Cells.TryGetValue(adress, out result);
            return result;
        }

        public MicroCell GetCellByHeader(string rowHeader, string Header)
        {
            int rowHeaderInt;

            if (!int.TryParse(rowHeader, out rowHeaderInt))
                throw new ArgumentException(string.Format("Row number {0} not digit", rowHeader));

            return GetCellByHeader(rowHeaderInt, Header);
        }


        public MicroCell GetCellByHeader(int rowHeaderInt, string Header)
        {
            var headers = this.HeadersDictionary;
            var targetHeader = headers.FirstOrDefault(h => h.Value == Header);

            //TODO:Есть ощущение что это плохая проверка
            if (targetHeader.Key == null && targetHeader.Value == null)
                throw new ArgumentOutOfRangeException(string.Format("Header {0} not exist in list {1}", Header, this.Name));

            var rows = this.RowsInt;

            if (!rows.Contains(rowHeaderInt))
                throw new ArgumentOutOfRangeException(string.Format("Row {0} not exist in list {1}", rowHeaderInt, this.Name));

            var targetRow = (List<MicroCell>)CellPacksSortedByRowHash[rowHeaderInt];

            var result = targetRow.FirstOrDefault(c => c.Column == targetHeader.Key);

            if (result == null)
                throw new Exception(string.Format("Cell with header {0} not exist in row {1}", Header, rowHeaderInt));

            return result;
        }

        private Dictionary<string, string> GetHeadersDictionary()
        {
            var result = new Dictionary<string, string>();
            foreach (var column in this.ColumnsList)
            {
                result.Add(column, GetCell(HeaderRowValue.ToString(), column).ViewValue);
            }
            return result;
        }

        public override string ToString()
        {
            return Name;
        }

        public Dictionary<string, MicroCell> GetAllCells()
        {
            return this.Cells;
        }

        public void Dispose()
        {
            foreach (var c in Cells)
            {
                c.Value.Dispose();
            }

        }
    }
}
