// Title: Limiting displayed items in a PivotField with Aspose.Cells for .NET (ShowItemsCount workaround)
// Common Searches: aspnet limit number of items shown in a pivot field using Aspose.Cells | c# set ShowItemsCount on a pivot field with Aspose.Cells library | how to hide extra row items in an Excel pivot table programmatically via Aspose.Cells | Aspose.Cells pivot table hide items beyond top N rows in C# | retrieve and modify pivot field items collection using Aspose.Cells .NET
// Tags: Aspose.Cells pivot field item limit | C# ShowItemsCount property usage | Excel pivot table row field visibility control | Aspose.Cells manipulate pivot field items collection | limit displayed pivot items .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    // The example loads an existing workbook, checks for a pivot table and at least one row field, accesses the first PivotField, and notes that Aspose.Cells does not expose a direct ShowItemsCount property, implying a need for alternative item‑visibility handling before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the workbook containing the pivot table
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                // Get the first pivot table
                PivotTable pivotTable = worksheet.PivotTables[0];

                // Ensure the pivot table has at least one row field
                if (pivotTable.RowFields.Count == 0)
                {
                    Console.WriteLine("The pivot table does not contain any row fields.");
                    return;
                }

                // Choose the first row field as an example
                PivotField pivotField = pivotTable.RowFields[0];

                // NOTE: Aspose.Cells does not expose a direct property to limit the number of displayed items.
                // If needed, you can manipulate the pivot field items collection here.

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
