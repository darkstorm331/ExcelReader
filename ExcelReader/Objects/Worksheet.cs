using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader.Objects
{
	public class Worksheet
	{
		public string Name { get; set; }
		public DataTable Data { get; set; }
	}
}
