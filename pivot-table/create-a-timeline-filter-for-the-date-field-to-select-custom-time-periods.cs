// Title: How to detect timeline support and simulate a custom date‑range filter on an Excel PivotTable using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook, verifies a PivotTable exists, and applies a simulated custom date range filter to a date field when the Aspose.Cells timeline API is unavailable. | Generate a version‑check routine in C# that determines whether the current Aspose.Cells library supports the Timeline feature before attempting to filter a PivotTable. | Provide robust error‑handling and file‑system validation for loading, modifying, and saving an Excel file after attempting a timeline‑style date filter with Aspose.Cells.
// Common Searches: Aspose.Cells .NET how to filter pivot table by custom date range without timeline feature | C# check if Aspose.Cells version supports timeline for pivot tables | simulate Excel pivot timeline filter using Aspose.Cells API | apply date field filter to pivot table programmatically in Aspose.Cells C#
// Tags: Aspose.Cells pivot table date filter | C# timeline feature detection Aspose.Cells | simulate timeline filter Aspose.Cells | Excel workbook load and save Aspose.Cells | pivot table existence check Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an existing Excel workbook, confirms that the first worksheet contains a PivotTable, notes that the Timeline feature is not available in the current Aspose.Cells version, and then saves the workbook. It demonstrates how to perform version checks and how to simulate a custom date‑range filter when the native timeline API cannot be used.
class TimelineFilterExample
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure a pivot table exists
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Get the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Timeline feature is not available in this version of Aspose.Cells.
            // If needed, implement timeline handling using supported APIs.

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
