using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsImportCustomObjectsDemo
{
    // Define a custom data class whose properties will be mapped to worksheet columns
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public int Stock { get; set; }
        public DateTime ReleaseDate { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Prepare a collection of custom objects to import
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple", Price = 2.99m, Stock = 150, ReleaseDate = new DateTime(2023, 12, 31) },
                new Product { Name = "Orange", Price = 1.99m, Stock = 200, ReleaseDate = new DateTime(2024, 1, 15) },
                new Product { Name = "Banana", Price = 0.99m, Stock = 300, ReleaseDate = new DateTime(2024, 2, 10) }
            };

            // Specify the property names to map to columns (order matters)
            string[] propertyNames = { "Name", "Price", "Stock", "ReleaseDate" };

            // Import the custom objects starting at row 2 (index 1) and column 1 (index 0)
            // Parameters:
            //   products                - collection to import
            //   propertyNames           - columns mapping
            //   true                    - include property names as header row
            //   1                       - firstRow (row index 1 => second row)
            //   0                       - firstColumn (column index 0 => first column)
            //   products.Count          - number of rows to import
            //   true                    - insert rows if needed
            //   "yyyy-MM-dd"            - date format for DateTime values
            //   true                    - try to convert strings to numbers where possible
            int importedRows = worksheet.Cells.ImportCustomObjects(
                products,
                propertyNames,
                true,
                1,
                0,
                products.Count,
                true,
                "yyyy-MM-dd",
                true
            );

            Console.WriteLine($"Imported {importedRows} rows starting at cell A2.");

            // Save the workbook to an XLSX file
            workbook.Save("ImportedProducts.xlsx");
        }
    }
}