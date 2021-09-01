using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReader.Objects;

namespace ExcelReader.Factory
{
	class Xls : IExcel
	{
		public Workbook GetWorkbook()
		{
			throw new NotImplementedException();
		}

		public Worksheet GetWorksheet(string sheetName)
		{
			throw new NotImplementedException();
		}

		public void Load(string filepath)
		{
			throw new NotImplementedException();
		}
	}
}
