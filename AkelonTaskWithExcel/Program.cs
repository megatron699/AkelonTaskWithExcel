using AkelonTaskWithExcel.Services;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Office2016.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Irony.Parsing;
using System.Reflection.Metadata;

namespace AkelonTaskWithExcel
{
	internal class Program
	{
		static void Main(string[] args)
		{
			try
			{
				var workbookService = new WorkbookService();

				Console.Write("Укажите путь к файлу: ");
				var path = Console.ReadLine();
				var workbook = workbookService.GetWorkbook(path);


				while (true)
				{
					Console.WriteLine("Введите цифру и нажмите Enter:\r\n" +
						"1. По наименованию товара вывести информацию о клиентах, заказавших этот товар, " +
						"с указанием информации по количеству товара, цене и дате заказа\r\n" +
						"2. Изменение контактного лица клиента.\r\n" +
						"3. Определить золотого клиента за указанный год\r\n" +
						"4. Определить золотого клиента за указанный месяц\r\n" +
						"0. Выход");

					var productTable = workbookService.GetTable(workbook, "Товары");
					var requestTable = workbookService.GetTable(workbook, "Заявки");
					var clientTable = workbookService.GetTable(workbook, "Клиенты");

					var menuButtonKey = Console.ReadLine();

					switch (menuButtonKey)
					{
						case "1":
							HandleByProduct(productTable, clientTable, requestTable);
							break;
						case "2":
							HandleByCompany(clientTable, workbook);
							break;
						case "3":
							HandleByYear(requestTable, clientTable);
							break;
						case "4":
							HandleByMonth(requestTable, clientTable);
							break;
						case "0":
							Exit();
							break;
						default:
							Console.WriteLine("Введена неизвестная команда");
							break;
					}
				}
			}
			catch(Exception ex)
			{
				Console.WriteLine(ex.Message);
				Console.ReadKey();
			}
		}

		private static void HandleByMonth(IXLTable requestTable, IXLTable clientTable)
		{
			Console.Write("Введите месяц в формате MM: ");
			var inputMonth = DateTime.ParseExact(Console.ReadLine(), "MM", null).Month;

			var maxRequests = requestTable.DataRange.Rows()
				.Where(x => x.Field("Дата размещения").Value.GetDateTime().Month == inputMonth)
				.GroupBy(x => x.Field("Код клиента").Value.GetNumber())
				.Max(t => t.Count());

			var goldClientsIds = requestTable.DataRange.Rows()
				.Where(x => x.Field("Дата размещения").Value.GetDateTime().Month == inputMonth)
				.GroupBy(x => x.Field("Код клиента").Value.GetNumber())
				.Where(x => x.Count() == maxRequests);

			foreach (var goldClientId in goldClientsIds)
			{
				var goldClient = clientTable.DataRange.Rows().First(x => x.Field("Код клиента").Value.GetNumber() == goldClientId.Key);

				Console.WriteLine($"Клиент {goldClient.Field("Наименование организации").Value.GetText()} совершил больше всего покупок за указанный месяц");
			}
		}

		private static void HandleByYear(IXLTable requestTable, IXLTable clientTable)
		{
			Console.Write("Введите год в формате yyyy: ");
			var inputYear = DateTime.ParseExact(Console.ReadLine(), "yyyy", null).Year; //TODO Create Exception

			var maxRequests = requestTable.DataRange.Rows()
				.Where(x => x.Field("Дата размещения").Value.GetDateTime().Year == inputYear)
				.GroupBy(x => x.Field("Код клиента").Value.GetNumber())
				.Max(t => t.Count());

			var goldClientsIds = requestTable.DataRange.Rows()
				.Where(x => x.Field("Дата размещения").Value.GetDateTime().Year == inputYear)
				.GroupBy(x => x.Field("Код клиента").Value.GetNumber())
				.Where(x => x.Count() == maxRequests);

			foreach (var id in goldClientsIds)
			{
				var goldClient = clientTable.DataRange.Rows().First(x => x.Field("Код клиента").Value.GetNumber() == id.Key);

				Console.WriteLine($"Клиент {goldClient.Field("Наименование организации").Value.GetText()} совершил больше всего покупок за указанный год");
			}
		}

		private static void HandleByProduct(IXLTable productTable, IXLTable clientTable, IXLTable requestTable)
		{
			var task1Service = new Task1Service();
			Console.Write("Введите название товара: ");
			var productName = Console.ReadLine();
			var productId = task1Service.GetProductIdByProductName(productTable, productName);
			if (productId == -1) //TODO Create Exception
			{
				Console.WriteLine("Данный товар не найден");
				return;
			}
			var productUnitPrice = task1Service.GetProductUnitPriceByProductId(productTable, productId);

			var infoes = task1Service.GetClientData(clientTable, requestTable, productId);

			task1Service.Print(infoes, productUnitPrice);
		}

		private static void HandleByCompany(IXLTable clientTable, IXLWorkbook workbook)
		{
			var task2Service = new Task2Service();
			Console.Write("Введите название организации: ");
			var companyName = Console.ReadLine();

			var client = task2Service.GetClientByCompanyName(clientTable, companyName);

			if (client == null) //TODO Create Exception
			{
				Console.WriteLine("Данный клиент не найден");
				return;
			}

			Console.Write("Введите новое ФИО: ");
			var newContactPerson = Console.ReadLine();

			task2Service.EditContactPerson(client, newContactPerson);
			workbook.Save();

			task2Service.PrintChanges(clientTable, companyName);
		}


		private static void Exit()
		{
			Environment.Exit(-1);
		}
	}
}