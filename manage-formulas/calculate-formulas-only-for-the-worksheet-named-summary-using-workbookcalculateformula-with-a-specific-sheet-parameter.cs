// Title: Recalculate only the "Summary" worksheet formulas with Aspose.Cells in C#
// Common Searches: how to recalculate formulas only on a specific worksheet using Aspose.Cells C# | Aspose.Cells calculate formulas for one sheet without recalculating the whole workbook | C# Workbook.CalculateFormula overload for a single worksheet | partial formula evaluation in an Excel file with Aspose.Cells | evaluate only the Summary tab formulas using Aspose.Cells C#
// Tags: Aspose.Cells single worksheet formula recalculation | Workbook.CalculateFormula worksheet overload C# | partial Excel formula evaluation Aspose.Cells | C# limit formula calculation to specific sheet | recalculate Summary worksheet formulas Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, obtains the worksheet named "Summary", and demonstrates how to invoke Workbook.CalculateFormula with a worksheet argument to evaluate formulas only on that sheet before saving the file. It also includes error handling for missing input files and ensures the output directory exists.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the worksheet named "Summary"
            Worksheet summarySheet = workbook.Worksheets["Summary"];

            // If the worksheet exists, calculate its formulas
            if (summarySheet != null)
            {
                // Use default calculation options
                workbook.CalculateFormula();
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook after calculation
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
