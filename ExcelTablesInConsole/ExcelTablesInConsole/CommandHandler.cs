using ClosedXML.Excel;

namespace ExcelTablesInConsole
{
    internal class CommandHandler
    {
        private ExcelService _excelService=new ExcelService();

        public void MenuService()
        {
            Console.WriteLine("Введите путь до файла с данными:");
            string filePath = Console.ReadLine();

            using (var workbook = new XLWorkbook(filePath))
            {
                bool _isWork = true;

                while (_isWork)
                {
                    Console.WriteLine();
                    Console.WriteLine("Выберите номер команды:");
                    Console.WriteLine("1. Поиск клиентов по наименованию товара");
                    Console.WriteLine("2. Изменение контактного лица клиента");
                    Console.WriteLine("3. Определение золотого клиента");
                    Console.WriteLine("4. Выйти из программы");

                    ConsoleKeyInfo key = Console.ReadKey(true);

                    switch (key.Key)
                    {
                        case ConsoleKey.NumPad1:
                            SearchCustomersByProductName(workbook);
                            break;
                        case ConsoleKey.NumPad2:
                            EditCustomerContact(workbook);
                            break;
                        case ConsoleKey.NumPad3:
                            FindGoldenCustomer(workbook);
                            break;
                        case ConsoleKey.NumPad4:
                            _isWork = false;
                            return;
                        default:
                            Console.WriteLine("Неверная команда. Попробуйте еще раз.");
                            break;
                    }

                    Console.ReadKey();
                    Console.Clear();
                }
            }
        }

        private void SearchCustomersByProductName(XLWorkbook workbook)
        {
            Console.WriteLine("Введите наименование товара:");
            string productName = Console.ReadLine();
            _excelService.SearchCustomersByProductName(workbook, productName);
        }

        private void EditCustomerContact(XLWorkbook workbook)
        {
            Console.WriteLine("Введите имя клиента:");
            string customerName = Console.ReadLine();

            Console.Write("\nВведите новое контактное лицо (ФИО): ");
            string newContact = Console.ReadLine();

            Console.Write("\nВведите название организации: ");
            string newCompany = Console.ReadLine();

            _excelService.EditCustomerContact(workbook, customerName, newContact, newCompany);
        }

        private void FindGoldenCustomer(XLWorkbook workbook)
        {
            Console.Write("Введите год: ");
            string yearInput = Console.ReadLine();

            Console.Write("Введите номер месяца: ");
            string monthInput = Console.ReadLine();

            _excelService.FindGoldenCustomer(workbook, yearInput, monthInput);
        }
    }
}
