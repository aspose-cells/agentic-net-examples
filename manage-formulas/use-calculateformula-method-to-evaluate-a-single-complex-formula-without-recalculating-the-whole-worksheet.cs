// Title: Use Worksheet.CalculateFormula to evaluate a single complex array formula in Aspose.Cells for .NET without full worksheet recalculation
// Common Searches: C# Aspose.Cells calculate a single SUM(IF) formula without recalculating the whole sheet | How to evaluate an array formula programmatically using Worksheet.CalculateFormula in .NET | Aspose.Cells evaluate complex Excel formula and write result to a specific cell | CalculateFormula method example for conditional sum in C# | Avoid full worksheet recalculation when computing a SUM(IF) range in Aspose.Cells
// Tags: Worksheet.CalculateFormula array formula evaluation | conditional sum calculation Aspose.Cells | single formula evaluation .NET Excel library | write calculation result to cell using Aspose.Cells | prevent full worksheet recalculation Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, defines a complex SUM(IF) array formula, evaluates it with Worksheet.CalculateFormula to obtain the result without triggering a full worksheet recalculation, writes the result into cell B1, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the complex formula to evaluate (example: sum of values in A1:A10 that are greater than 5)
            string formula = "SUM(IF(A1:A10>5, A1:A10, 0))";

            // Evaluate the formula using the first worksheet's context
            Worksheet sheet = workbook.Worksheets[0];
            object result = sheet.CalculateFormula(formula);

            // Display the result
            Console.WriteLine("Formula result: " + result);

            // Write the result back to a cell (B1) in the worksheet
            Cell resultCell = sheet.Cells["B1"];
            resultCell.PutValue(result);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the new value
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
