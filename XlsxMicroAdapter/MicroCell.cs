using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlsxMicroAdapter
{
	public class MicroCell
	{
		public string Row { get; set; }

		public string Column { get; set; }

		public string ViewValue { get; set; }

		public string FormulaValue { get; set; }


		public MicroCell(string row, string column, string viewValue = "", string formula = "")
		{
			this.Row = row;
			this.Column = column;
			this.ViewValue = viewValue;
			this.FormulaValue = formula;

		}

		public MicroCell()
		{
			
		}
		
		public override string ToString()
		{
			return string.Concat(this.Column, this.Row);
		}
	}
}
