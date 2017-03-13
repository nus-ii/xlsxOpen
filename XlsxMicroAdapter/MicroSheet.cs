using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlsxMicroAdapter
{
	public class MicroSheet
	{
		public string Name { get; set; }

		private List<MicroCell> Cells { get; set; }

		public bool Visible { get; set; }

		private List<string> RowsList;
		private List<string> ColumnsList;

		public List<string> Rows
		{
			get
			{
				FixColumnListAndRowList();
				return RowsList;
			}
		}

		public List<string> Columns
		{
			get
			{
				FixColumnListAndRowList();
				return ColumnsList;
			}
		}

		public Dictionary<string, string> HeadersDictionary
		{
			get { return GetHeadersDictionary(); }
		}

		public MicroSheet(string name = "", bool visible = true)
		{
			this.Name = name;
			this.Cells = new List<MicroCell>();
			this.Visible = visible;
			this.ColumnsList=new List<string>();
			this.RowsList=new List<string>();

		}

		public void AddCell(MicroCell newCell)
		{
			this.Cells.Add(newCell);
			//FixColumnListAndRowList();
		}

		private void FixColumnListAndRowList()
		{
			foreach (var cell in this.Cells)
			{
				if (!ColumnsList.Contains(cell.Column))
					ColumnsList.Add(cell.Column);

				if (!RowsList.Contains(cell.Row))
					RowsList.Add(cell.Row);
			}

		}

		public void AddCellList(List<MicroCell> CellList)
		{
			this.Cells.AddRange(CellList);
		    //FixColumnListAndRowList();
		}

		public List<MicroCell> GetCellsWhereRow(string row)
		{
			return this.Cells.Where(c => c.Row == row).ToList();
		}

		public List<MicroCell> GetCellsWhereColumn(string column)
		{
			return this.Cells.Where(c => c.Column == column).ToList();
		}

		public List<string> GetColumns()
		{
			var result = new List<string>();
			foreach (var cell in Cells)
			{
				if (result.Count == 0 || !result.Contains(cell.Column))
					result.Add(cell.Column);
			}
			return result;
		}

		public List<string> GetRows()
		{
			var result = new List<string>();
			foreach (var cell in Cells)
			{
				if (result.Count == 0 || !result.Contains(cell.Row))
					result.Add(cell.Row);
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
			return Cells.FirstOrDefault(c => c.Row == row && c.Column == column);
		}

		private Dictionary<string, string> GetHeadersDictionary()
		{
			var result = new Dictionary<string, string>();
			foreach (var column in this.ColumnsList)
			{
				result.Add(column, GetCell("1", column).ViewValue);
			}
			return result;
		}

		public override string ToString()
		{
			return Name;
		}
	}
}
