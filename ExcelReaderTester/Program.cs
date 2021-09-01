using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReader.Factory;

namespace ExcelReaderTester
{
	class Program
	{
		static void Main(string[] args)
		{
			IExcel test = new Xlsx();


			test.Load("C:\\Users\\micha\\Dropbox\\My PC (DESKTOP-QR8R89P)\\Desktop\\testfile.xlsx");
			Console.ReadKey();
		}
	}
}
