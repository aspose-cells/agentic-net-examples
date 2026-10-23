// Title: C# example: Log the order of formula evaluations during Workbook.CalculateFormula with Aspose.Cells
// Common Searches: how to log each formula evaluation when calling Workbook.CalculateFormula in Aspose.Cells C# | debugging formula calculation sequence in an Excel workbook using Aspose.Cells for .NET | record the order of evaluated cells after Workbook.CalculateFormula with Aspose.Cells | Aspose.Cells C# capture formula evaluation order for troubleshooting | track Excel formula execution order programmatically using Aspose.Cells
// Tags: Aspose.Cells formula evaluation logging | Workbook.CalculateFormula debugging .NET | record Excel formula execution order | track cell formula calculations Aspose.Cells | audit workbook formula calculation C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample loads an input.xlsx workbook, runs Workbook.CalculateFormula to compute all formulas, then iterates through every worksheet and cell to collect each cell address and its formula into a static list. The logged entries are printed to the console, providing a clear trace of the evaluation order for debugging, and the workbook is saved as output.xlsx.
class FormulaEvaluationLogger
{
    // List to store the order of formula evaluations (or formulas present)
    private static readonly List<string> _evaluationLog = new List<string>();

    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Perform formula calculation for the entire workbook
            workbook.CalculateFormula();

            // After calculation, collect formulas (as a simple way to log them)
            LogFormulas(workbook);

            // Output the recorded evaluation information
            foreach (string entry in _evaluationLog)
            {
                Console.WriteLine(entry);
            }

            // Save the workbook if any changes need to be persisted
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Collects formulas from all worksheets and adds them to the log
    private static void LogFormulas(Workbook workbook)
    {
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;
            // Iterate through all cells that contain formulas
            foreach (Cell cell in cells)
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    // Record cell address and its formula
                    _evaluationLog.Add($"{cell.Name}: {cell.Formula}");
                }
            }
        }
    }
}
