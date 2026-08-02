using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsImportCustomObjectsDemo
{
    // Sample data class
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Prepare sample data
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 2.99m, Stock = 150, Date = new DateTime(2023, 12, 31) },
                new Product { Name = "Orange", Price = 1.99m, Stock = 200, Date = new DateTime(2024, 1, 15) }
            };

            // Mapping dictionary: original property name -> desired column header
            Dictionary<string, string> columnMapping = new Dictionary<string, string>
            {
                { "Name",  "Product Name" },
                { "Price", "Unit Price"   },
                { "Stock", "Inventory"    },
                { "Date",  "Sale Date"    }
            };

            // Property names to import (must be the original property names)
            string[] propertyNames = new string[columnMapping.Count];
            int idx = 0;
            foreach (var kvp in columnMapping)
                propertyNames[idx++] = kvp.Key;

            // Import the custom objects; include property names as the first row
            int importedRows = cells.ImportCustomObjects(
                products,               // list
                propertyNames,          // propertyNames
                true,                   // isPropertyNameShown (adds header row)
                0,                      // firstRow
                0,                      // firstColumn
                products.Count,         // rowNumber
                true,                   // insertRows
                "yyyy-MM-dd",           // dateFormatString
                true                    // convertStringToNumber
            );

            // Rename the header cells according to the mapping dictionary
            // Header row is at firstRow (0) because we set isPropertyNameShown = true
            for (int col = 0; col < propertyNames.Length; col++)
            {
                string originalProp = propertyNames[col];
                if (columnMapping.TryGetValue(originalProp, out string newHeader))
                {
                    cells[0, col].PutValue(newHeader);
                }
            }

            // Save the workbook
            workbook.Save("ImportCustomObjectsWithRenamedColumns.xlsx");
        }
    }
}