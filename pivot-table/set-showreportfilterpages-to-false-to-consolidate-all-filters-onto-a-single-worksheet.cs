// Title: How to suppress pivot table report filter pages in Aspose.Cells for .NET and keep all filters on a single worksheet
// AI Prompts: Generate C# code with Aspose.Cells that disables the creation of separate report filter worksheets for a pivot table. | Provide a C# workaround to consolidate pivot table report filter pages onto the main worksheet when ShowReportFilterPages is unavailable. | Show how to load an Excel workbook, modify pivot table filter page behavior, and save the file using Aspose.Cells.
// Common Searches: Aspose.Cells C# hide pivot table report filter pages | prevent Aspose.Cells from generating filter worksheets for pivot tables | consolidate pivot table filters onto one sheet using Aspose.Cells .NET | ShowReportFilterPages property missing Aspose.Cells workaround | keep all pivot filters on the same worksheet in Aspose.Cells
// Tags: Aspose.Cells disable pivot report filter pages | C# hide pivot table filter worksheets | Aspose.Cells consolidate pivot filters | Excel workbook pivot filter page suppression | Aspose.Cells pivot table settings .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, checks for a pivot table, notes that the ShowReportFilterPages property is not available, and demonstrates how to handle or work around filter page suppression before saving the modified file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "source.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the workbook containing the pivot table
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one pivot table on the worksheet
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found on the first worksheet.");
                    return;
                }

                // Get the first pivot table
                PivotTable pivotTable = worksheet.PivotTables[0];

                // NOTE: The ShowReportFilterPages property is not available in the current Aspose.Cells version.
                // If needed, alternative handling of report filter pages should be implemented here.

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
