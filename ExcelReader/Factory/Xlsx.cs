using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReader.Objects;
using System.IO;

namespace ExcelReader.Factory
{
	public class Xlsx : IExcel
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
			try
			{
				if (File.Exists(filepath))
				{
					using (FileStream fs = File.OpenRead(filepath))
					{
						byte[] b = new byte[1024];
						UTF8Encoding temp = new UTF8Encoding(true);
						while (fs.Read(b, 0, b.Length) > 0)
						{
							Console.WriteLine(temp.GetString(b));
						}
					}
				}
				else
				{

				}
			}
			catch
			{

			}
		}
	}
}
