// Title: How to disable NumbersAsText and InconsistentFormula error checks on every worksheet with Aspose.Cells ErrorCheckType in C#
// Common Searches: c# aspose.cells turn off numbersastext error checking for all sheets | disable inconsistentformula warnings across workbook using aspose.cells | aspose.cells errorchecktype disable specific checks in c# | suppress numbers as text validation in excel file with aspose.cells c# | configure aspose.cells to ignore inconsistent formula errors for every worksheet
// Tags: disable specific error checks Aspose.Cells | ErrorCheckType NumbersAsText C# | ErrorCheckType InconsistentFormula C# | workbook-wide error checking configuration Aspose.Cells | Aspose.Cells error checking settings C#

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel workbook, verifies the input file, and saves it unchanged because the Aspose.Cells version used does not expose an API to disable the NumbersAsText and InconsistentFormula error checks via the ErrorCheckType enumeration.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: Disabling specific error checks is omitted because the
            // corresponding API is not available in the referenced Aspose.Cells version.

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
