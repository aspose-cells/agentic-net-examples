// Title: Calculate all formulas in an Excel workbook with Aspose.Cells for .NET – partial range A1:D20 not supported
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, verifies the file exists, runs Workbook.CalculateFormula to recalculate all formulas, and saves the workbook. | Explain why Aspose.Cells cannot recalculate formulas only in a defined range and suggest alternative approaches for evaluating a specific cell block.
// Common Searches: Aspose.Cells calculate formulas for a specific cell range in C# | Does Aspose.Cells support partial formula recalculation .NET | How to evaluate only A1:D20 formulas using Aspose.Cells | Workaround to recalculate selected cells with Aspose.Cells | Workbook.CalculateFormula limitations in Aspose.Cells .NET
// Tags: Workbook.CalculateFormula full recalculation | Aspose.Cells partial formula evaluation limitation | C# evaluate Excel formulas Aspose.Cells | Excel workbook formula calculation Aspose.Cells | selective range calculation workaround Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads 'input.xlsx' into an Aspose.Cells Workbook, checks that the file exists, calls Workbook.CalculateFormula (which recomputes every formula because partial recalculation is not supported), and saves the result to 'output.xlsx' with basic error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Recalculate all formulas in the workbook (partial recalculation not supported in this version)
            workbook.CalculateFormula();

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
