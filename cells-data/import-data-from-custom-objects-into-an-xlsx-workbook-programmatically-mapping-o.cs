using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCustomObjectImport
{
    // Define a custom data class whose properties will be mapped to worksheet cells
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public int Stock { get; set; }
        public DateTime ReleaseDate { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();

            // 2. Get the first worksheet where data will be imported
            Worksheet worksheet = workbook.Worksheets[0];

            // 3. Prepare a collection of custom objects
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 2.99m, Stock = 150, ReleaseDate = new DateTime(2023, 12, 31) },
                new Product { Name = "Orange", Price = 1.99m, Stock = 200, ReleaseDate = new DateTime(2024, 1, 15) },
                new Product { Name = "Banana", Price = 0.99m, Stock = 300, ReleaseDate = new DateTime(2024, 2, 10) }
            };

            // 4. Specify which properties to import and the order
            string[] propertyNames = { "Name", "Price", "Stock", "ReleaseDate" };

            // 5. Import the custom objects into the worksheet
            //    - Show property names as header row
            //    - Start at cell A1 (row 0, column 0)
            //    - Insert rows if needed
            //    - Use a date format for DateTime values
            //    - Convert numeric strings to numbers when possible
            int importedRows = worksheet.Cells.ImportCustomObjects(
                products,                 // ICollection list
                propertyNames,            // string[] propertyNames
                true,                     // bool isPropertyNameShown
                0,                        // int firstRow
                0,                        // int firstColumn
                products.Count,           // int rowNumber
                true,                     // bool insertRows
                "yyyy-MM-dd",             // string dateFormatString
                true);                    // bool convertStringToNumber

            Console.WriteLine($"Successfully imported {importedRows} rows (including header).");

            // 6. Auto‑fit columns for better readability (optional)
            worksheet.AutoFitColumns();

            // 7. Save the workbook to an XLSX file (lifecycle save rule)
            workbook.Save("ProductsReport.xlsx");

            // Dispose the workbook to release resources
            workbook.Dispose();
        }
    }
}