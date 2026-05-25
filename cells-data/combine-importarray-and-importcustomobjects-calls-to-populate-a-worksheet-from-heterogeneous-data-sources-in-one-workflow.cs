using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCombinedImportDemo
{
    // Sample custom object to be imported
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
            Cells cells = worksheet.Cells;

            // -------------------------------------------------
            // 1. Import a simple string array as column headers
            // -------------------------------------------------
            string[] headers = { "Product Name", "Unit Price", "Units In Stock", "Release Date" };
            // ImportArray(string[] array, int firstRow, int firstColumn, bool isVertical)
            // Place headers in the first row (row index 0), starting from column A (column index 0), horizontally
            cells.ImportArray(headers, 0, 0, false);

            // -------------------------------------------------
            // 2. Prepare a collection of custom objects (Product)
            // -------------------------------------------------
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple", Price = 2.99m, Stock = 150, ReleaseDate = new DateTime(2023, 12, 31) },
                new Product { Name = "Orange", Price = 1.99m, Stock = 200, ReleaseDate = new DateTime(2024, 1, 15) },
                new Product { Name = "Banana", Price = 0.99m, Stock = 300, ReleaseDate = new DateTime(2024, 2, 10) }
            };

            // Property names to import (order must match the headers)
            string[] propertyNames = { "Name", "Price", "Stock", "ReleaseDate" };

            // -------------------------------------------------
            // 3. Import the custom objects below the header row
            // -------------------------------------------------
            // ImportCustomObjects(ICollection list, string[] propertyNames, bool isPropertyNameShown,
            //                    int firstRow, int firstColumn, int rowNumber,
            //                    bool insertRows, string dateFormatString, bool convertStringToNumber)
            // firstRow = 1 (second row) because row 0 holds headers
            // firstColumn = 0 (column A)
            // rowNumber = products.Count
            // insertRows = true (add rows if needed)
            // dateFormatString = "yyyy-MM-dd"
            // convertStringToNumber = true (convert numeric strings)
            cells.ImportCustomObjects(
                products,
                propertyNames,
                false,          // do NOT repeat property names (already added as headers)
                1,              // start importing at row index 1
                0,              // start at column A
                products.Count,
                true,           // insert rows if necessary
                "yyyy-MM-dd",   // date format for DateTime values
                true            // try to convert strings to numbers
            );

            // -------------------------------------------------
            // 4. Save the workbook
            // -------------------------------------------------
            workbook.Save("CombinedImportDemo.xlsx");
        }
    }
}