using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AkelonTaskWithExcel.Services
{
	internal class Task1Service
	{
		

		internal int GetProductIdByProductName(IXLTable productTable, string productName)
		{
			var product = productTable.DataRange.Rows()
				.FirstOrDefault(productFromTable => productFromTable.Field("Наименование").Value.ToString() == productName);
			if(product == null)
			{
				return -1;
			}
			return (int) product.Field("Код товара").Value
				.GetNumber();
		}

		internal int GetProductUnitPriceByProductId(IXLTable productTable, int productId)
		{
			return (int)productTable.DataRange.Rows()
				.First(productFromTable => (int)productFromTable.Field("Код товара").Value.GetNumber() == productId)
				.Field("Цена товара за единицу").Value.GetNumber();
		}

		internal dynamic GetClientData(IXLTable clientTable, IXLTable requestTable, int productId)
		{
			return clientTable.DataRange.Rows()
				.Join(
					requestTable.DataRange.Rows(),
					clientFromTable => (int)clientFromTable.Field("Код клиента").Value.GetNumber(),
					requestFromTable => (int)requestFromTable.Field("Код клиента").Value.GetNumber(),
					(clientFromTable, requestFromTable) => new
					{
						companyName = clientFromTable.Field("Наименование организации").Value.GetText(),
						address = clientFromTable.Field("Адрес").Value.GetText(),
						contactPerson = clientFromTable.Field("Контактное лицо (ФИО)").Value.GetText(),
						requestedAmount = (int)requestFromTable.Field("Требуемое количество").Value.GetNumber(),
						published = requestFromTable.Field("Дата размещения").Value.GetDateTime(),
						productId = (int)requestFromTable.Field("Код товара").Value.GetNumber()
					}
				)
				.Where(requestFromJoin => requestFromJoin.productId == productId);
		}

		internal void Print(dynamic clientsInfo, int productUnitPrice)
		{
			foreach (var clientInfo in clientsInfo)
			{
				Console.WriteLine($"Организация {clientInfo.companyName} по адресу {clientInfo.address} (ФИО контактного лица: {clientInfo.contactPerson}) " +
					$"заказала данный товар в количестве {clientInfo.requestedAmount} на сумму {clientInfo.requestedAmount * productUnitPrice} {clientInfo.published}");
			}
		}		
	}
}

