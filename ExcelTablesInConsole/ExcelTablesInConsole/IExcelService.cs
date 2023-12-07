using ClosedXML.Excel;

namespace ExcelTablesInConsole
{
    internal interface IExcelService
    {
        public string ProductsWorksheet { get; }
        public string EntryWorksheet { get; }
        public string CustomersWorksheet { get; }

        public void SearchCustomersByProductName(XLWorkbook workbook, string productName);
        public void EditCustomerContact(XLWorkbook workbook, string customerName, string newContact, string newCompany);
        public void FindGoldenCustomer(XLWorkbook workbook, string yearInput, string monthInput);
    }
}
