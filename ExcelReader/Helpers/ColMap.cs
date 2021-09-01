using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.Helpers
{
	public static class ColMap
	{
		public static int AlphaToNum(string alpha)
		{
			int[] digits = new int[alpha.Length];
			for (int i = 0; i < alpha.Length; ++i)
			{
				digits[i] = Convert.ToInt32(alpha[i]) - 64;
			}

			int mul = 1; 
			int res = 0;

			for (int pos = digits.Length - 1; pos >= 0; --pos)
			{
				res += digits[pos] * mul;
				mul *= 26;
			}

			return res;
		}

		public static string NumToAlpha(int num)
		{
			string columnName = "";

			while (num > 0)
			{
				int modulo = (num - 1) % 26;
				columnName = Convert.ToChar('A' + modulo) + columnName;
				num = (num - modulo) / 26;
			}

			return columnName;
		}
	}
}
