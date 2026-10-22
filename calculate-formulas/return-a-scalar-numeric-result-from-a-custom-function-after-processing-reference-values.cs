// Title: Compute a single numeric result from a SUM‑multiply formula and retrieve it in C# using Aspose.Cells
// Common Searches: Aspose.Cells C# calculate SUM of a range and multiply by a constant | How to read a double value from a cell after evaluating formulas with Aspose.Cells | Using Workbook.CalculateFormula to get scalar results in .NET | Saving an Excel workbook after formula calculation with Aspose.Cells | Create output folder automatically before saving Excel file in C#
// Tags: sum range multiplication Aspose.Cells | read cell double after calculation C# | Workbook.CalculateFormula usage | save workbook Aspose.Cells .NET | ensure directory exists before saving Excel

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, fills cells A1‑A3 with numbers, assigns the formula SUM(A1:A3)*2 to B1, calculates all formulas, reads the resulting double from B1, ensures the target directory exists, and saves the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Populate sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Use built‑in functions to achieve the same result as the intended custom function
            // Formula: (SUM of A1:A3) * 2
            sheet.Cells["B1"].Formula = "SUM(A1:A3)*2";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Retrieve the scalar result from the cell
            double result = sheet.Cells["B1"].DoubleValue;
            Console.WriteLine("Result: " + result); // Expected output: 120

            // Define output file path
            string outputPath = "Result.xlsx";

            // Ensure the directory exists before saving
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
