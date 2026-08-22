// Title: How to combine ImportArray and ImportCustomObjects to import headers, product objects, and vertical totals in a single Aspose.Cells .NET worksheet
// AI Prompts: Use ImportArray to write a header row, then call ImportCustomObjects to load a List<Product>, and finally import a vertical numeric array with ImportArray to add totals in the same sheet. | Build a single import routine that sequentially adds headers, imports custom objects, and appends a column of totals, handling row offsets automatically.
// Common Searches: Aspose.Cells C# import header row and list of objects in one worksheet | ImportCustomObjects after ImportArray example Aspose.Cells | Add vertical totals column after importing objects with Aspose.Cells | Combine array and custom object imports in an Aspose.Cells workbook | C# heterogeneous data import workflow using Aspose.Cells
// Tags: ImportArray header row Aspose.Cells | ImportCustomObjects product list C# | ImportArray vertical totals column | heterogeneous data import Aspose.Cells | combined array and object import workflow

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

namespace AsposeCellsDemo
{
    // The example creates a workbook, adds a header row with ImportArray, imports a List<Product> using ImportCustomObjects (skipping duplicate headers), then appends a vertical numeric array of totals with ImportArray, and saves the result as an .xlsx file.
    public class HeterogeneousImportDemo
    {
        // Custom data class to be imported via ImportCustomObjects
        public class Product
        {
            public string Name { get; set; } = string.Empty;
            public decimal Price { get; set; }
            public int Stock { get; set; }
            public DateTime ReleaseDate { get; set; }
        }

        public static void Run()
        {
            try
            {
                // Create a new workbook and obtain the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // 1. Import a header row using ImportArray (horizontal)
                string[] headers = new string[] { "Product", "Price", "Stock", "ReleaseDate" };
                // firstRow = 0, firstColumn = 0, isVertical = false (horizontal)
                cells.ImportArray(headers, 0, 0, false);

                // 2. Prepare a list of custom objects (heterogeneous source)
                List<Product> products = new List<Product>
                {
                    new Product { Name = "Apple", Price = 2.99m, Stock = 150, ReleaseDate = new DateTime(2023, 12, 31) },
                    new Product { Name = "Orange", Price = 1.99m, Stock = 200, ReleaseDate = new DateTime(2024, 1, 15) }
                };

                // Define the properties to import; can be null to import all
                string[] propertyNames = { "Name", "Price", "Stock", "ReleaseDate" };

                // 3. Import the custom objects starting just below the header row
                // Use ImportCustomObjects(ICollection, string[], bool, int, int, int, bool, string, bool)
                int importedRows = cells.ImportCustomObjects(
                    products,          // list
                    propertyNames,     // propertyNames
                    false,             // isPropertyNameShown (headers already added)
                    1,                 // firstRow (row index after header)
                    0,                 // firstColumn
                    products.Count,    // rowNumber (number of rows to import)
                    true,              // insertRows (add rows if needed)
                    "yyyy-MM-dd",      // dateFormatString
                    true               // convertStringToNumber
                );

                // 4. Import a numeric array vertically after the custom object rows
                double[] totals = new double[] { 1000.0, 2000.0 };
                int startRowForTotals = 1 + importedRows; // position after the last product row
                // ImportArray(double[], int, int, bool) – vertical placement in column E (index 4)
                cells.ImportArray(totals, startRowForTotals, 4, true);

                // Determine output file path
                string outputPath = "HeterogeneousImportDemo.xlsx";

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            HeterogeneousImportDemo.Run();
        }
    }
}
