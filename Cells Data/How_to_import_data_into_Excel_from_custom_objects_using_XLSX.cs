using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCustomObjectImportDemo
{
    // Define a custom class whose instances will be imported into Excel
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
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // 2. Prepare a collection of custom objects
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 2.99m, Stock = 150, ReleaseDate = new DateTime(2023, 12, 31) },
                new Product { Name = "Orange", Price = 1.99m, Stock = 200, ReleaseDate = new DateTime(2024, 1, 15) },
                new Product { Name = "Banana", Price = 0.99m, Stock = 300, ReleaseDate = new DateTime(2024, 2, 10) }
            };

            // 3. Specify which properties to import (order matters)
            string[] propertyNames = { "Name", "Price", "Stock", "ReleaseDate" };

            // 4. Import the custom objects into the worksheet
            //    Parameters:
            //    - collection of objects
            //    - property names array
            //    - show property names as header row (true)
            //    - start row index (0 = first row)
            //    - start column index (0 = first column)
            //    - number of rows to import (list count)
            //    - insert rows if needed (true)
            //    - date format string for DateTime values
            //    - convert string values to numbers when possible (true)
            int importedRows = worksheet.Cells.ImportCustomObjects(
                products,
                propertyNames,
                true,
                0,
                0,
                products.Count,
                true,
                "yyyy-MM-dd",
                true);

            Console.WriteLine($"Successfully imported {importedRows} rows.");

            // 5. Save the workbook in XLSX format
            workbook.Save("CustomObjectsImportDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}