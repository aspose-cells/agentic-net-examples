using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsListObjectDemo
{
    // Sample data class
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public int Quantity { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 1.20m, Quantity = 50 },
                new Product { Name = "Banana", Price = 0.80m, Quantity = 80 },
                new Product { Name = "Orange", Price = 1.00m, Quantity = 60 }
            };

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Import the list as a table (including headers)
            ImportTableOptions importOptions = new ImportTableOptions
            {
                IsFieldNameShown = true,   // include property names as header row
                InsertRows = true,
                ConvertNumericData = true
            };
            // Import starts at cell A1 (row 0, column 0)
            sheet.Cells.ImportCustomObjects(products, 0, 0, importOptions);

            // Determine the range of the imported data
            int startRow = 0;               // A1
            int startColumn = 0;            // column A
            int endRow = products.Count;    // header + data rows => row index = count (since header occupies row 0)
            int endColumn = 2;              // three columns: Name, Price, Quantity (0,1,2)

            // Add a ListObject (Excel table) over the imported range
            // The range includes the header row, so hasHeaders = true
            int tableIndex = sheet.ListObjects.Add(startRow, startColumn, endRow, endColumn, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "ProductTable";
            table.TableStyleType = TableStyleType.TableStyleMedium9;

            // Save the workbook as XLSX
            workbook.Save("ProductsListObject.xlsx", SaveFormat.Xlsx);
        }
    }
}