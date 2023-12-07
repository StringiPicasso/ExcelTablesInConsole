using ClosedXML.Excel;

namespace ExcelTablesInConsole
{
    internal class ExcelService: IExcelService
    {
        public string productsWorksheet => "Товары";
        public string entryWorksheet => "Заявки";
        public string customersWorksheet => "Клиенты";

        public void SearchCustomersByProductName(XLWorkbook workbook, string productName)
        {
            var productsRows = GetNeccessaryRangeRow(workbook, productsWorksheet, "B", productName);
            var productCode = productsRows.Cell("A").Value;
            var priceProduct = productsRows.Cell("D").GetDouble();

            var entryRows = GetNeccessaryRangeRow(workbook, entryWorksheet, "B", productCode.ToString());
            var customerCode = entryRows.Cell("C").Value;
            var countProduct = entryRows.Cell("E").GetDouble();
            var dataPlacement = entryRows.Cell("F").Value;
            var allprice = priceProduct * countProduct;

            var cusomersRows = GetNeccessaryRangeRow(workbook, customersWorksheet, "A", customerCode.ToString());
            var customerName = cusomersRows.Cell("B").Value;

            Console.WriteLine($"\nКлиент: {customerName}\nДата заказа: {dataPlacement}\nКоличество: {customerCode}\nЦена: {allprice}");
        }

        public void EditCustomerContact(XLWorkbook workbook, string customerName, string newContact, string newCompany)
        {
            var customerRows = GetNeccessaryRangeRow(workbook, customersWorksheet, "D", customerName);

            Console.WriteLine("Обновление данных контактного лица: " + customerRows.Cell("D").Value);

            customerRows.Cell("B").Value = newCompany;
            customerRows.Cell("D").Value = newContact;

            Console.WriteLine("\n\nКонтактное лицо успешно изменено.");
            Console.WriteLine($"\nРезультат изменений:\nОрганизация: {customerRows.Cell("B").Value}\nКонтактное лицо: {customerRows.Cell("D").Value}\nАдрес: {customerRows.Cell("C").Value}");
        }

        public void FindGoldenCustomer(XLWorkbook workbook,string yearInput, string monthInput)
        {
            DateTime orderDate;

            //Check whether the input are int
            if (!int.TryParse(yearInput, out int year) || !int.TryParse(monthInput, out int month))
            {
                Console.WriteLine("Неверно введены год или месяц.");
                return;
            }

            var worksheet = workbook.Worksheet(entryWorksheet);
            var orders = worksheet.RangeUsed().RowsUsed();

            var customerOrders = new Dictionary<string, int>();

            //Using the orderDate, look for matches by year and month and determine the customer's code
            foreach (var order in orders)
            {
                if (DateTime.TryParse(order.Cell("F").Value.ToString(), out orderDate))
                {
                    if (orderDate.Year == year && orderDate.Month == month)
                    {
                        var codeCustomer = order.Cell("C").Value;

                        var custonerRow= GetNeccessaryRangeRow(workbook, customersWorksheet, "A", codeCustomer.ToString());

                        string customerName = custonerRow.Cell("B").Value.ToString();

                        //Adding suitable customers to the Dictionary
                        if (customerOrders.ContainsKey(customerName))
                        {
                            customerOrders[customerName]++;
                        }
                        else
                        {
                            customerOrders[customerName] = 1;
                        }
                    }
                }
            }

            if (customerOrders.Count > 0)
            {
                //Return the first one in the sorted Dictionary by Value, determine the appropriate customer
                string goldenCustomer = customerOrders.OrderByDescending(x => x.Value).First().Key;
                Console.WriteLine("Золотой клиент за указанный год и месяц: " + goldenCustomer);
            }
            else
            {
                Console.WriteLine("Нет данных о заказах за указанный год и месяц.");
            }
        }

        private IXLRangeRow GetNeccessaryRangeRow(XLWorkbook workbook, string worksheet, string nameColumn, string compareName)
        {
            var xLWorksheet = workbook.Worksheet(worksheet);
            var rangeRows = xLWorksheet.RangeUsed().RowsUsed();

            foreach (var item in rangeRows)
            {
                if (item.Cell(nameColumn).Value.ToString() == compareName)
                {
                    return item;
                }
            }

            return null;
        }
    }
}
