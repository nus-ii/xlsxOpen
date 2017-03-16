using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlsxMicroAdapter
{
	public class DataCheckInfo
	{
		public MicroCell SourceTopLeft;
		public MicroCell SourceBottomRight;
		public string SourceSheetName;
		public MicroCell Target;

		public DataCheckInfo()
		{
			
		}

		public DataCheckInfo(MicroCell top,MicroCell bottom,string source, MicroCell target)
		{
			this.SourceTopLeft = top;
			this.SourceBottomRight = bottom;
			this.SourceSheetName = source;
			this.Target = target;
		}
	}
}
