using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlsxMicroAdapter
{
	public class XlsxHelper
	{
		private static string[] alph = new string[]{
			"A","B","C","D","E","F","G","H","I","J","K","L","M",
			"N","O","P","Q","R","S","T","U","V","W","X","Y","Z"};

		public static string GetColumnLetter(int i)
		{
			if(i>700)
				throw new ArgumentException("");

			string result = "";
			
			if (i <= alph.Length)
			{
				result = alph.ElementAt(i - 1);
			}
			else
			{
				int n = i / alph.Length;
				int y = i % alph.Length;

				if (y == 0)
				{
					y = alph.Length;
					n = n - 1;
				}

				result = string.Concat(alph.ElementAt(n - 1), alph.ElementAt(y - 1));

			}

			return result;
		}
	}
}
