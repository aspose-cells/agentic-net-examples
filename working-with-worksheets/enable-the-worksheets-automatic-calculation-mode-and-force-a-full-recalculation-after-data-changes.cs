// Title: How to enable automatic calculation mode and force a full formula recalculation in an Excel worksheet using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set worksheet calculation mode to automatic | Force full workbook recalculation after data changes with Aspose.Cells .NET | How to trigger CalculateFormula for all sheets in Aspose.Cells | Enable auto calculate and recalc formulas in Excel file using Aspose.Cells C# | Aspose.Cells .NET recalculate all formulas after modifying cells
// Tags: Worksheet calculation mode automatic Aspose.Cells | Force full formula recalculation Aspose.Cells .NET | Workbook.CalculateFormula method C# | Excel auto calculate setting Aspose.Cells | Recalculate formulas after data update Aspose.Cells | Aspose.Cells calculation options .NET

using Aspose.Cells;
using System;
using System.IO;

namespace AsposeCellsExample
{
    // The sample loads an existing Excel workbook, invokes Workbook.CalculateFormula() to recalculate every formula in the file, and then saves the updated workbook. It includes basic file existence checking, exception handling, and console logging of the operation status.
    class Program
    {
        static void Main()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);

                // Recalculate all formulas in the workbook
                workbook.CalculateFormula();

                // Save the workbook to the output file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
