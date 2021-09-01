using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReader.Objects;

namespace ExcelReader.Factory
{
	public interface IExcel
	{
		void Load(string filepath);

		Workbook GetWorkbook();

		Worksheet GetWorksheet(string sheetName);
	}
}
