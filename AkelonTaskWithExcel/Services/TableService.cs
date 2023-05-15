using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AkelonTaskWithExcel.Services
{
	internal class WorkbookService
	{
		internal XLWorkbook GetWorkbook(string path)
		{
			try
			{
				var workbook = new XLWorkbook(path);
				return workbook;
			}
			catch (ArgumentException)
			{
				throw new ArgumentException("Файл по указанному пути не найден");
			}
		}
		internal IXLTable GetTable(XLWorkbook workbook, string worksheetName)
		{
			var worksheet = workbook.Worksheet(worksheetName);
			var firstRowUsed = worksheet.FirstRowUsed();
			var row = firstRowUsed.RowUsed();
			var firstPossibleAddress = worksheet.Row(row.RowNumber()).FirstCell().Address;
			var lastPossibleAddress = worksheet.LastCellUsed().Address;
			var range = worksheet.Range(firstPossibleAddress, lastPossibleAddress).RangeUsed();
			return range.AsTable();
		}
	}
}
