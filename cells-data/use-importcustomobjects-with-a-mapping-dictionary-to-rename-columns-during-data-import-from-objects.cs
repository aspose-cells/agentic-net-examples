using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsImportCustomObjectsDemo
{
    // Sample custom object
    public class Product
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public int Stock { get; set; }
        public DateTime Date { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // 2. Prepare sample data
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 2.99m, Stock = 150, Date = new DateTime(2023,12,31) },
                new Product { Name = "Orange", Price = 1.99m, Stock = 200, Date = new DateTime(2024, 1, 15) }
            };

            // 3. Mapping dictionary: original property name -> desired column header
            Dictionary<string, string> columnRenameMap = new Dictionary<string, string>
            {
                { "Name",  "Product Name" },
                { "Price", "Unit Price"   },
                { "Stock", "Quantity"     },
                { "Date",  "Release Date" }
            };

            // 4. Property names to import (must match object properties)
            string[] propertyNames = { "Name", "Price", "Stock", "Date" };

            // 5. Import the custom objects with property names shown as header row
            cells.ImportCustomObjects(
                products,               // list
                propertyNames,          // propertyNames
                true,                   // isPropertyNameShown (header row)
                0,                      // firstRow
                0,                      // firstColumn
                products.Count,         // rowNumber
                true,                   // insertRows
                "yyyy-MM-dd",           // dateFormatString
                true                    // convertStringToNumber
            );

            // 6. Rename header cells according to the mapping dictionary
            // Header row is at index 0 (firstRow)
            for (int col = 0; col < propertyNames.Length; col++)
            {
                string originalName = propertyNames[col];
                if (columnRenameMap.TryGetValue(originalName, out string newName))
                {
                    cells[0, col].PutValue(newName);
                }
            }

            // 7. Save the workbook
            workbook.Save("ImportCustomObjectsWithRenamedColumns.xlsx");
        }
    }
}