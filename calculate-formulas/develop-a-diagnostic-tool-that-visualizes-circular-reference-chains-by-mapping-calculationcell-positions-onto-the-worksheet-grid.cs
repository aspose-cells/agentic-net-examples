// Title: C# utility using Aspose.Cells to detect circular reference errors during workbook formula calculation and generate a diagnostic file
// Common Searches: how to detect circular reference errors with Aspose.Cells when calculating formulas in C# | Aspose.Cells C# catch formula calculation exception for circular references | save a diagnostic Excel file after formula calculation failure using Aspose.Cells | C# program to log circular reference messages from Aspose.Cells workbook processing | load workbook and force formula calculation with Aspose.Cells C# example
// Tags: Aspose.Cells calculate formula circular reference detection | C# exception handling Aspose.Cells formula calculation | diagnostic workbook generation Aspose.Cells | load and save Excel workbook Aspose.Cells C# | log circular reference errors Aspose.Cells

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace CircularReferenceDiagnostic
{
    // The sample loads an Excel workbook with Aspose.Cells, forces a full formula calculation to surface any circular reference exceptions, catches and logs the error message, and then saves the workbook (or a diagnostic copy) to a designated output path.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputPath = @"C:\Input\WorkbookWithCircularRefs.xlsx";
            string outputPath = @"C:\Output\WorkbookWithCircularRefs_Diagnosed.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            Workbook workbook;
            try
            {
                // Load the workbook
                workbook = new Workbook(inputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            try
            {
                // Force calculation of all formulas.
                // Aspose.Cells will throw an exception if a circular reference is detected.
                workbook.CalculateFormula();
            }
            catch (Exception ex)
            {
                // Log the exception message (circular reference or other calculation errors)
                Console.WriteLine("Error during formula calculation: " + ex.Message);
                // Optional: attempt to highlight cells that may be involved.
                // Since specific exception details are unavailable, this step is omitted.
            }

            try
            {
                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the (potentially modified) workbook
                workbook.Save(outputPath);
                Console.WriteLine("Diagnostic workbook saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
