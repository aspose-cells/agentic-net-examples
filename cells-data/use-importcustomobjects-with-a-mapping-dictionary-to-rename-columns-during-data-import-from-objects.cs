// Title: Import a C# List of objects into Excel and rename column headers using a dictionary with Aspose.Cells ImportCustomObjects
// AI Prompts: Generate C# code that uses Aspose.Cells ImportCustomObjects to load a List<Product> into a worksheet, then replace the default property name headers with custom titles from a Dictionary<string,string>. | Show how to apply a specific date format and enable string‑to‑number conversion while importing custom objects with Aspose.Cells, and then save the workbook. | Provide a snippet that inserts rows, maps object properties to columns, renames the header row based on a mapping dictionary, and writes the file to disk.
// Common Searches: how to rename Excel column headers after importing a C# object list with Aspose.Cells | Aspose.Cells ImportCustomObjects custom header names dictionary example | C# import list of objects to Excel with date format and string to number conversion using Aspose.Cells | mapping object property names to custom Excel column titles with Aspose.Cells ImportCustomObjects
// Tags: ImportCustomObjects header mapping Aspose.Cells | custom column titles from dictionary C# | date format string conversion Aspose.Cells | insert rows during object import Aspose.Cells | rename Excel headers after ImportCustomObjects

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsImportCustomObjectsDemo
{
    // The program creates a workbook, imports a List<Product> into the first worksheet using ImportCustomObjects, replaces the automatically generated property name headers with custom titles defined in a dictionary, applies a date format and string‑to‑number conversion, inserts rows as needed, and saves the file as ImportCustomObjectsWithRenamedColumns.xlsx.
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

            // Sample data
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple",  Price = 2.99m, Stock = 150, Date = new DateTime(2023, 12, 31) },
                new Product { Name = "Orange", Price = 1.99m, Stock = 200, Date = new DateTime(2024, 1, 15) }
            };

            // Mapping from original property names to desired column headers
            Dictionary<string, string> columnMap = new Dictionary<string, string>
            {
                { "Name",  "Product Name" },
                { "Price", "Unit Price" },
                { "Stock", "Available Stock" },
                { "Date",  "Release Date" }
            };

            // Property names to import (must match the object's property names)
            string[] propertyNames = { "Name", "Price", "Stock", "Date" };

            // Import the custom objects; property names will be written to the first row
            int importedRows = cells.ImportCustomObjects(
                products,
                propertyNames,
                true,               // isPropertyNameShown
                0,                  // firstRow
                0,                  // firstColumn
                products.Count,     // rowNumber
                true,               // insertRows
                "yyyy-MM-dd",       // dateFormatString
                true                // convertStringToNumber
            );

            // Rename the header cells according to the mapping dictionary
            for (int col = 0; col < propertyNames.Length; col++)
            {
                string originalName = propertyNames[col];
                if (columnMap.TryGetValue(originalName, out string newHeader))
                {
                    cells[0, col].PutValue(newHeader);
                }
            }

            // Save the workbook
            workbook.Save("ImportCustomObjectsWithRenamedColumns.xlsx");
        }
    }
}
