// Title: Refresh all slicers and recalculate formulas in an Excel workbook using Aspose.Cells for .NET
// Common Searches: how to programmatically refresh slicers in an Excel file with Aspose.Cells C# | Aspose.Cells recalculate workbook formulas after slicer refresh in .NET | C# code to refresh all slicers and run CalculateFormula on a workbook | update slicer filters and recalc dependent formulas using Aspose.Cells API | refresh slicer objects before saving workbook with Aspose.Cells for .NET
// Tags: slicer.Refresh method Aspose.Cells | Workbook.CalculateFormula after slicer refresh | programmatic slicer refresh C# | recalculate formulas after slicer update .NET | Aspose.Cells Excel slicer handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// The example loads an Excel workbook, iterates through each worksheet to call Refresh on every slicer, invokes Workbook.CalculateFormula to update all dependent formulas, ensures the output directory exists, and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Refresh all slicers to apply the current filter selections
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Slicer slicer in sheet.Slicers)
                {
                    slicer.Refresh();
                }
            }

            // Recalculate all formulas after the slicer refresh
            workbook.CalculateFormula();

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
