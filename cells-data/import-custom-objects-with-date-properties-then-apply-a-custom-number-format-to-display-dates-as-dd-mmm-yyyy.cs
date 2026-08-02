using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCustomObjectImportDemo
{
    // Define a custom class with a DateTime property
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
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Prepare a collection of custom objects
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple", Price = 2.99m, Stock = 150, Date = new DateTime(2023, 12, 31) },
                new Product { Name = "Orange", Price = 1.99m, Stock = 200, Date = new DateTime(2024, 1, 15) }
            };

            // Specify the property names to import (including the date property)
            string[] propertyNames = { "Name", "Price", "Stock", "Date" };

            // Import the custom objects.
            // The dateFormatString parameter sets the desired display format for DateTime cells.
            worksheet.Cells.ImportCustomObjects(
                products,               // ICollection list
                propertyNames,          // string[] propertyNames
                true,                   // isPropertyNameShown (include header row)
                0,                      // firstRow (zero‑based)
                0,                      // firstColumn (zero‑based)
                products.Count,         // rowNumber (number of rows to import)
                true,                   // insertRows (add rows if needed)
                "dd-MMM-yyyy",          // dateFormatString (custom date format)
                true                    // convertStringToNumber (attempt numeric conversion)
            );

            // Save the workbook (lifecycle: save)
            workbook.Save("ImportCustomObjectsWithCustomDateFormat.xlsx");
        }
    }
}