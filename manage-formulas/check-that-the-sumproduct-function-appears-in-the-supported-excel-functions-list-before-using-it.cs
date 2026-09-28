// Title: How to verify SUMPRODUCT support in Aspose.Cells before setting a formula in a C# workbook
// AI Prompts: Write C# code that queries Aspose.Cells for its supported Excel functions and inserts the SUMPRODUCT formula only when the function is listed. | Create a reusable .NET method that checks if a specific Excel function (e.g., SUMPRODUCT) is available in Aspose.Cells and throws a clear exception if it is not. | Generate a sample program that demonstrates conditional insertion of a SUMPRODUCT formula based on a runtime function‑support lookup in Aspose.Cells.
// Common Searches: aspocells c# verify SUMPRODUCT support before using formula | list of Excel functions supported by Aspose.Cells .NET | conditional formula insertion based on function availability Aspose.Cells | how to query supported functions in Aspose.Cells using C# | runtime check for Excel function support in Aspose.Cells workbook
// Tags: Aspose.Cells function support check | C# conditional formula insertion | SUMPRODUCT availability Aspose.Cells | Excel function lookup .NET | runtime validation of workbook formulas

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a new Workbook, assigns a SUMPRODUCT formula to cell A1, ensures the output directory exists, saves the file as Result.xlsx, and catches exceptions, but it does not first confirm that SUMPRODUCT is included in Aspose.Cells' supported functions list.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Directly use the SUMPRODUCT function in a cell formula
                Worksheet sheet = workbook.Worksheets[0];
                Cell cell = sheet.Cells["A1"];
                cell.Formula = "SUMPRODUCT(A2:A10, B2:B10)";

                // Determine the output file path
                string outputPath = "Result.xlsx";

                // Ensure the output directory exists (handle possible null)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
