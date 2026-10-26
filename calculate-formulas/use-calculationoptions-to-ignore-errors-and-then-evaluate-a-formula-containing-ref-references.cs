// Title: Ignore #REF! errors while evaluating Excel formulas using Aspose.Cells CalculationOptions in C#
// Common Searches: Aspose.Cells C# calculate formula with #REF! reference ignored | set CalculationOptions.IgnoreError to true for Excel formula evaluation in .NET | evaluate Excel formula that points to a missing sheet using Aspose.Cells | how to prevent exceptions when calculating formulas with invalid references in Aspose.Cells | use Workbook.CalculateFormula with error‑handling options in C#
// Tags: CalculationOptions.IgnoreError property | Aspose.Cells formula calculation without errors | C# evaluate formula referencing non‑existent sheet | Workbook.CalculateFormula error suppression | handling #REF! errors in Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, inserts a formula that references a missing sheet (producing #REF!), configures CalculationOptions.IgnoreError to true, evaluates all formulas, retrieves the resulting value (0), and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a formula that references a non‑existent sheet (will cause #REF! error)
            sheet.Cells["A1"].Formula = "=Sheet2!A1";

            // Set calculation options to ignore errors during evaluation
            CalculationOptions calcOptions = new CalculationOptions
            {
                IgnoreError = true
            };

            // Evaluate all formulas using the specified options
            workbook.CalculateFormula(calcOptions);

            // Retrieve the evaluated value (error is ignored, result will be 0)
            object result = sheet.Cells["A1"].Value;
            Console.WriteLine($"Evaluated value in A1: {result}");

            // Define output file path
            string outputPath = "Result.xlsx";

            // Save the workbook (overwrite if it already exists)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
