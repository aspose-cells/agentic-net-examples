// Title: Skip merged‑cell validation when importing a C# object collection into an Excel worksheet with Aspose.Cells
// AI Prompts: Write C# code that uses Aspose.Cells ImportCustomObjects with ImportTableOptions.CheckMergedCells set to false to overwrite merged cells. | Show how to import a List<T> into a specific start cell while disabling merged‑cell checks and suppressing the header row. | Provide a complete example that merges a range, then imports objects into the same sheet without triggering merged‑cell errors.
// Common Searches: Aspose.Cells import custom objects without merged cell validation C# | How to disable merged cell checking when using ImportCustomObjects | Import a list of C# objects into Excel while ignoring merged ranges | Set ImportTableOptions.CheckMergedCells false example Aspose.Cells | Write data over merged cells in Excel using Aspose.Cells C#
// Tags: ImportCustomObjects with CheckMergedCells disabled | skip merged cell validation Aspose.Cells | import C# collection into Excel worksheet | write over merged cells using Aspose.Cells | InsertRows option for ImportCustomObjects

using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsImportSkipMergedCells
{
    // Sample class to be imported
    // The sample creates a workbook, merges cells B2:C3, and then imports a List<Product> into the worksheet starting at A1. ImportTableOptions are configured with CheckMergedCells = false, IsFieldNameShown = false, and InsertRows = true, allowing data to be written over the merged range. The workbook is saved as ImportSkipMergedCells.xlsx.
    public class Product
    {
        public string Name { get; set; } = string.Empty;
        public double Price { get; set; }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Create a merged cell range for demonstration (cells B2:C3)
                cells.Merge(1, 1, 2, 2); // Merge B2:C3
                cells[1, 1].PutValue("MergedCell");

                // Prepare sample data to import
                List<Product> products = new List<Product>
                {
                    new Product { Name = "Apple",  Price = 1.20 },
                    new Product { Name = "Banana", Price = 0.80 },
                    new Product { Name = "Cherry", Price = 2.50 }
                };

                // Configure import options
                ImportTableOptions importOptions = new ImportTableOptions
                {
                    // Skip checking merged cells so data will be written over them
                    CheckMergedCells = false,
                    // Do not import property names as a header row
                    IsFieldNameShown = false,
                    // Insert rows if needed (optional)
                    InsertRows = true
                };

                // Import the custom objects starting at cell A1 (row 0, column 0)
                cells.ImportCustomObjects((ICollection)products, 0, 0, importOptions);

                // Save the workbook
                workbook.Save("ImportSkipMergedCells.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
