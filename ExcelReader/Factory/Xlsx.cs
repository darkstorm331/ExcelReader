using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelReader.Objects;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;

namespace ExcelReader.Factory
{
	public class Xlsx : IExcel
	{
		public XDocument App { get; set; }
		public XDocument Core { get; set; }
		public XDocument Custom { get; set; }
		public XDocument Workbook { get; set; }
		public XDocument ContentTypes { get; set; }
		public XDocument SharedStrings { get; set; }
		public XDocument Styles { get; set; }
		public Dictionary<string, XDocument> Sheets { get; set; }

		public Workbook WholeWorkbook { get; set; }

		public Workbook GetWorkbook()
		{
			//Set up Workbook

			return WholeWorkbook;
		}

		public Worksheet GetWorksheet(string sheetName)
		{
			foreach (Worksheet ws in WholeWorkbook.Worksheets)
			{
				if (ws.Name == sheetName) 
				{
					return ws;
				}
			}

			return null;
		}

		public void Load(string filepath)
		{
			try
			{
				Sheets = new Dictionary<string, XDocument>();

				if (File.Exists(filepath))
				{
					using (FileStream fs = new FileStream(filepath, FileMode.Open, FileAccess.Read)) 
					{
						//Read in the first 4 bytes to check this is an Xlsx file
						//Beginning 4 bytes should be 50 4B 03 04 (https://www.filesignatures.net/index.php?search=XLSX&mode=EXT)
						byte[] headerBytes = new byte[4];
						fs.Read(headerBytes, 0, 4);

						if (BitConverter.ToString(headerBytes) == "50-4B-03-04")
						{					
							using (ZipArchive za = ZipFile.OpenRead(filepath)) 
							{
								foreach (ZipArchiveEntry entry in za.Entries) 
								{
									using (Stream s = entry.Open()) 
									{
										using (var sr = new StreamReader(s, Encoding.UTF8))
										{
											string content = sr.ReadToEnd();

											switch (entry.FullName) 
											{
												case "xl/workbook.xml":
													Workbook = XDocument.Parse(content);
													break;

												case "xl/styles.xml":
													Styles = XDocument.Parse(content);
													break;

												case "xl/sharedStrings.xml":
													SharedStrings = XDocument.Parse(content);
													break;

												case "docProps/core.xml":
													Core = XDocument.Parse(content);
													break;

												case "docProps/app.xml":
													App = XDocument.Parse(content);
													break;

												case "docProps/custom.xml":
													Custom = XDocument.Parse(content);
													break;

												case "[Content_Types].xml":
													ContentTypes = XDocument.Parse(content);
													break;

												default:
													if (entry.FullName.StartsWith("xl/worksheets")) { Sheets.Add(Path.GetFileNameWithoutExtension(entry.Name), XDocument.Parse(content)); }
													break;
											}
										}
									}
								}
							}
						}
						else
						{
							throw new Exception("Signature of file does not match valid Excel signature");
						}
					}
				}
				else
				{
					throw new Exception("File does not exist or could not be accessed");
				}
			}
			catch (Exception err)
			{
				Console.Error.WriteLine(err.GetBaseException().Message);
				throw;
			}
		}
	}
}
