// Title: Updating a named range and refreshing dependent formulas using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change value in a named range and recalculate workbook formulas | C# code to modify Excel named range and trigger Workbook.CalculateFormula in Aspose.Cells | How to refresh formulas after updating a named range with Aspose.Cells for .NET | Recalculate dependent cells after editing a named range using Aspose.Cells in C# | Aspose.Cells example: update MyRange and recalc formulas programmatically
// Tags: named range update with Aspose.Cells | Workbook.CalculateFormula method | refresh dependent formulas in Excel using Aspose.Cells | programmatic Excel named range modification C# | Aspose.Cells formula recalculation after cell edit

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, accesses the named range "MyRange", sets its first cell to 100, calls Workbook.CalculateFormula to recompute all formulas that depend on the changed cell, and saves the result to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range
            Name namedRange = workbook.Worksheets.Names["MyRange"];
            if (namedRange == null)
            {
                Console.WriteLine("Named range 'MyRange' not found.");
                return;
            }

            // Get the actual cell range that the name refers to
            Aspose.Cells.Range cellRange = namedRange.GetRange();

            // Update the first cell of the named range
            cellRange[0, 0].PutValue(100);

            // Recalculate all formulas
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
