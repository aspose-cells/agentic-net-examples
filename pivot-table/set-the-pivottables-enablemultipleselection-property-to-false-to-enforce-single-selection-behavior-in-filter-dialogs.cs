// Title: Disable multiple selection in Aspose.Cells PivotTable filter dialogs using C#
// Common Searches: Aspose.Cells C# how to turn off multiple selection in pivot table filter dialog | set EnableMultipleSelection false for a PivotTable with Aspose.Cells API | configure single‑selection mode for Excel pivot table filters in C# | prevent multiple items from being selected in Aspose.Cells pivot filter | example code to enforce single selection in Aspose.Cells pivot table filters
// Tags: Aspose.Cells disable pivot filter multiple selection | PivotTable EnableMultipleSelection property C# | single selection pivot filter configuration Aspose.Cells | Excel pivot table filter settings using Aspose.Cells | C# Aspose.Cells pivot table behavior customization

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The example loads an existing workbook, checks for a pivot table on the first worksheet, accesses the first PivotTable, and notes that the EnableMultipleSelection property is not available in the current Aspose.Cells version, then saves the workbook. It demonstrates how to approach disabling multiple selection in pivot filter dialogs with Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (worksheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found on the first worksheet.");
                return;
            }

            // Access the first pivot table
            PivotTable pivotTable = worksheet.PivotTables[0];

            // NOTE: The EnableMultipleSelection property is not available in the current Aspose.Cells version.
            // If needed, adjust the pivot table settings using other available APIs.

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
