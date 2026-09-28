// Title: Set workbook to Automatic calculation mode, recalculate all formulas, and read updated cell values using Aspose.Cells for .NET
// AI Prompts: Load an Excel file, assign Workbook.Settings.CalculationMode = CalculationModeType.Automatic, invoke workbook.CalculateFormula(), then output the value of cell A1. | Switch the workbook to automatic calculation, trigger a full formula refresh, and verify the result of a dependent cell before saving. | Programmatically enable automatic evaluation, recalculate the entire sheet, and retrieve the refreshed value of a target cell to confirm the update.
// Common Searches: C# Aspose.Cells how to turn on automatic formula evaluation and refresh all cells | recalculate workbook formulas after changing calculation settings with Aspose.Cells | retrieve updated value of a cell after calling CalculateFormula in Aspose.Cells .NET | save modified workbook after enabling automatic calculation using Aspose.Cells
// Tags: Workbook.Settings.CalculationMode Automatic Aspose.Cells | recalculate all formulas using Aspose.Cells API | read cell value after formula evaluation Aspose.Cells | save updated workbook to XLSX Aspose.Cells | handle workbook load errors Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel workbook, switches the calculation mode to Automatic, recalculates every formula, reads the refreshed value of cell A1, and saves the updated file, with proper exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Example: read the value of cell A1 after recalculation
            var worksheet = workbook.Worksheets[0];
            var cell = worksheet.Cells["A1"];
            Console.WriteLine($"A1 value after recalculation: {cell.Value}");

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
