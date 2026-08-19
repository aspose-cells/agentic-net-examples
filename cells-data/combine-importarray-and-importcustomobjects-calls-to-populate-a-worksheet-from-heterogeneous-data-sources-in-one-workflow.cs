// Title: Combine ImportArray and ImportCustomObjects in Aspose.Cells .NET to Load Headers, Objects, and Adjustments
// Description: Demonstrates how to create a workbook, add column headers with ImportArray, import a List<Product> using ImportCustomObjects (skipping duplicate headers, inserting rows, applying a date format, and converting strings to numbers), then import a vertical integer array of stock adjustments, and finally save the file as an Excel workbook.
// Keywords: Aspose.Cells | ImportArray C# | ImportCustomObjects C# | worksheet data import | custom object list import | vertical array import | date formatting Aspose.Cells | convert string to number | product catalog Excel | stock adjustment column | C# Excel automation
// Common Searches: Aspose.Cells import list of objects and array in same sheet | ImportArray together with ImportCustomObjects example | C# import vertical integer array after custom objects Aspose.Cells | how to add headers then import objects in Aspose.Cells | combine heterogeneous data sources in Aspose.Cells workbook
// Developer Intent: The developer needs a single workflow that adds column headers, imports a collection of custom objects, and then appends a vertical numeric array—all within one worksheet using Aspose.Cells for .NET.
// Use Cases: Generate a product catalog where static headers, dynamic product records, and separate stock‑adjustment values are populated automatically. | Create a financial report that mixes predefined titles, a list of transaction objects, and a column of correction amounts without manual cell addressing. | Build a data‑migration script that consolidates CSV‑style arrays and object collections into a formatted Excel sheet in one pass.
// AI Prompts: Write C# code with Aspose.Cells that imports a string array as column headers, then a List<T> of custom objects using ImportCustomObjects (with date format and number conversion), and finally a vertical integer array starting after the imported rows. | Show how to calculate the start row for a second ImportArray call after ImportCustomObjects has inserted rows, and include code to save the workbook. | Explain best practices for combining ImportArray and ImportCustomObjects to handle heterogeneous data sources in a single worksheet.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsCombinedImportDemo
{
    // Sample custom object to be imported via ImportCustomObjects
    // Demonstrates how to create a workbook, add column headers with ImportArray, import a List<Product> using ImportCustomObjects (skipping duplicate headers, inserting rows, applying a date format, and converting strings to numbers), then import a vertical integer array of stock adjustments, and finally save the file as an Excel workbook.
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

            // ------------------------------------------------------------
            // 1. Import a simple string array using ImportArray (horizontal)
            // ------------------------------------------------------------
            string[] headers = new string[] { "Product Name", "Price", "Stock", "Release Date" };
            // Import headers starting at cell A1 (row 0, column 0) horizontally
            cells.ImportArray(headers, 0, 0, false);

            // ------------------------------------------------------------
            // 2. Prepare a list of custom objects (heterogeneous data source)
            // ------------------------------------------------------------
            List<Product> products = new List<Product>
            {
                new Product { Name = "Apple", Price = 2.99m, Stock = 150, ReleaseDate = new DateTime(2023, 12, 31) },
                new Product { Name = "Orange", Price = 1.99m, Stock = 200, ReleaseDate = new DateTime(2024, 1, 15) },
                new Product { Name = "Banana", Price = 0.99m, Stock = 300, ReleaseDate = new DateTime(2024, 2, 10) }
            };

            // Define the property names to import (order matters)
            string[] propertyNames = { "Name", "Price", "Stock", "ReleaseDate" };

            // ------------------------------------------------------------
            // 3. Import the custom objects using ImportCustomObjects
            //    - Show property names in the first row (already added above, so set false)
            //    - Start importing data at row 1 (second row), column 0 (A column)
            //    - Insert rows if needed, use a date format, and convert strings to numbers
            // ------------------------------------------------------------
            int importedRows = cells.ImportCustomObjects(
                products,                // ICollection list
                propertyNames,           // string[] propertyNames
                false,                   // isPropertyNameShown (already added)
                1,                       // firstRow (row index where data starts)
                0,                       // firstColumn
                products.Count,          // rowNumber (number of rows to import)
                true,                    // insertRows
                "yyyy-MM-dd",            // dateFormatString
                true                     // convertStringToNumber
            );

            Console.WriteLine($"Imported {importedRows} product rows.");

            // ------------------------------------------------------------
            // 4. Import an integer array vertically below the custom objects
            // ------------------------------------------------------------
            int[] stockAdjustments = new int[] { -10, 5, -20 };
            // Determine the start row: headers (0) + data rows (importedRows) + 1 empty row
            int startRowForAdjustments = 1 + importedRows + 1;
            // Import vertically starting at column 4 (E column) to keep it separate
            cells.ImportArray(stockAdjustments, startRowForAdjustments, 4, true);

            // ------------------------------------------------------------
            // 5. Save the workbook
            // ------------------------------------------------------------
            workbook.Save("CombinedImportDemo.xlsx");
        }
    }
}
