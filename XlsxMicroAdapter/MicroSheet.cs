using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace XlsxMicroAdapter
{
	public class MicroSheet
	{
		public string Name { get; set; }

		private Dictionary<string,MicroCell> Cells { get; set; }

		public bool Visible { get; set; }

		private List<string> RowsList;
		private List<string> ColumnsList;

		public List<DataCheckInfo> CheckList;

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
			this.Cells = new Dictionary<string, MicroCell>();
			this.Visible = visible;
			this.ColumnsList=new List<string>();
			this.RowsList=new List<string>();
			this.CheckList=new List<DataCheckInfo>();

		}

		public void AddCell(MicroCell newCell)
		{
			this.Cells.Add(string.Concat(newCell.Column,newCell.Row),newCell);
		}

		private void FixColumnListAndRowList()
		{
			foreach (var cell in this.Cells)
			{
				if (!ColumnsList.Contains(cell.Value.Column))
					ColumnsList.Add(cell.Value.Column);

				if (!RowsList.Contains(cell.Value.Row))
					RowsList.Add(cell.Value.Row);
			}

		}

		public void AddCellList(List<MicroCell> CellList)
		{
			foreach (var cell in CellList)
			{
				this.AddCell(cell);
			}
		}

		public List<MicroCell> GetCellsWhereRow(string row)
		{
			List<MicroCell> result=new List<MicroCell>();
			var preResult=this.Cells.Where(c => c.Value.Row == row).ToList();

			foreach (var cell in preResult)
			{
				result.Add(cell.Value);
			}
			return result;

		}

		public List<MicroCell> GetCellsWhereColumn(string column)
		{
			List<MicroCell> result = new List<MicroCell>();
			var preResult=this.Cells.Where(c => c.Value.Column == column).ToList();
			foreach (var cell in preResult)
			{
				result.Add(cell.Value);
			}
			return result;
		}

		public List<string> GetColumns()
		{
			var result = new List<string>();
			foreach (var cell in Cells)
			{
				if (result.Count == 0 || !result.Contains(cell.Value.Column))
					result.Add(cell.Value.Column);
			}
			return result;
		}

		public List<string> GetRows()
		{
			var result = new List<string>();
			foreach (var cell in Cells)
			{
				if (result.Count == 0 || !result.Contains(cell.Value.Row))
					result.Add(cell.Value.Row);
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
			MicroCell result=new MicroCell();
			Cells.TryGetValue(adress, out result);
			return result;
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
