using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AkelonTaskWithExcel.Services
{
	internal class Task2Service
	{
		internal IXLTableRow GetClientByCompanyName(IXLTable clientTable, string companyName)
		{
			var client = clientTable.DataRange.Rows()
				.FirstOrDefault(clientFromFile => clientFromFile.Field("Наименование организации").Value.GetText() == companyName);
			if(client == null)
			{
				return null;
			}
			return client;
		}

		internal void EditContactPerson(IXLTableRow client, string newContactPerson)
		{
			client.Field("Контактное лицо (ФИО)").SetValue(newContactPerson);
			
		}

		internal void PrintChanges(IXLTable clientTable, string companyName)
		{
			Console.WriteLine($"ФИО контакного лица изменено на " +
						$"{clientTable.DataRange.Rows().First(x => x.Field("Наименование организации").Value.GetText() == companyName).Field("Контактное лицо (ФИО)").Value.GetText()}");
		}
	}
}
