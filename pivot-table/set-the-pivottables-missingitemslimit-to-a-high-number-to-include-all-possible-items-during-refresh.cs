// Title: Set a high MissingItemsLimit on an Aspose.Cells PivotTable in C# to include all items during refresh
// AI Prompts: Assign a large integer to PivotTable.MissingItemsLimit before invoking RefreshData with Aspose.Cells for .NET. | Update the C# sample to set MissingItemsLimit to Int32.MaxValue so the refreshed PivotTable retains every possible item.
// Common Searches: Aspose.Cells C# how to keep all items when refreshing a PivotTable | set MissingItemsLimit property to max value in Aspose.Cells PivotTable | include hidden items in PivotTable refresh using Aspose.Cells .NET | C# Aspose.Cells PivotTable missing items limit example | increase MissingItemsLimit to avoid item truncation in Aspose.Cells
// Tags: Aspose.Cells PivotTable MissingItemsLimit | C# set PivotTable MissingItemsLimit | Aspose.Cells refresh PivotTable all items | PivotTable include hidden items Aspose.Cells | Aspose.Cells .NET PivotTable data refresh configuration

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an Excel workbook, accesses the first PivotTable, sets its MissingItemsLimit to a high value to ensure all items are retained, refreshes the PivotTable data, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the PivotTable
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one PivotTable
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No PivotTables found in the worksheet.");
                return;
            }

            // Access the first PivotTable on the worksheet
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Refresh the PivotTable data using the correct API
            pivotTable.RefreshData();

            // Ensure output directory exists (if any)
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
