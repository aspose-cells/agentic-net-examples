using System;
using System.IO;
using Aspose.Cells;

namespace MyApp
{
    class ReplaceNamedRange
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Retrieve the named range by name
                const string rangeName = "MyRange";
                Name namedRange = workbook.Worksheets.Names[rangeName];

                // Verify the named range exists
                if (namedRange == null)
                {
                    Console.WriteLine($"Named range \"{rangeName}\" not found.");
                    return;
                }

                // Update the named range reference
                namedRange.RefersTo = "=Sheet1!$A$1:$A$10";

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
}