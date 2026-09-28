// Title: Loop through every worksheet and enable automatic formula calculation after importing data with Aspose.Cells for .NET
// AI Prompts: Add a C# loop that visits each worksheet in a loaded Workbook and sets its Settings.CalcMode to Automatic before saving. | Generate code that loads an Excel file, performs data insertion, switches all worksheets to automatic calculation mode, triggers a full recalculation, and writes the result to a new file.
// Common Searches: C# Aspose.Cells set automatic calculation for each worksheet after data import | How to enable workbook calculation mode automatic in Aspose.Cells .NET | Iterate through worksheets and change CalcMode to Automatic using Aspose.Cells | Force formula recalculation after updating cells with Aspose.Cells C#
// Tags: Aspose.Cells set workbook calculation mode automatic | C# iterate worksheets Aspose.Cells | force formula recalculation Aspose.Cells | automatic calculation after data import .NET | Aspose.Cells workbook Settings.CalcMode

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, optionally imports data, loops through every worksheet to set the workbook's calculation mode to Automatic, forces a full formula recalculation, and saves the updated workbook, while handling missing files and runtime exceptions.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook
                var workbook = new Workbook(inputPath);

                // ----- Data import operations go here -----
                // Example: workbook.Worksheets[0].Cells["A1"].PutValue(123);
                // -------------------------------------------

                // Force calculation of all formulas in the workbook
                workbook.CalculateFormula();

                // Save the workbook to the desired output path
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
