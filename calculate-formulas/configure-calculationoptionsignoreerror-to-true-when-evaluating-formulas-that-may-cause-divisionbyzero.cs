// Title: How to suppress #DIV/0! errors by setting CalculationOptions.IgnoreError in Aspose.Cells for .NET
// Common Searches: Aspose.Cells ignore division by zero during workbook.CalculateFormula in C# | set CalculationOptions.IgnoreError true to avoid #DIV/0! in Aspose.Cells | C# example for handling formula errors with Aspose.Cells CalculationOptions | prevent Excel #DIV/0! exception when recalculating formulas using Aspose.Cells | ignore formula errors while calculating workbook in Aspose.Cells .NET
// Tags: Aspose.Cells CalculationOptions IgnoreError | C# ignore #DIV/0! Aspose.Cells | formula error handling Aspose.Cells .NET | Workbook.CalculateFormula error suppression | division by zero handling Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// Creates an in‑memory workbook, inserts a formula that divides by zero, configures CalculationOptions.IgnoreError to true, recalculates all formulas so the cell returns 0 instead of an error, and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a formula that could cause division‑by‑zero
            sheet.Cells["A1"].Formula = "10/0";

            // Configure calculation options to ignore errors such as division‑by‑zero
            CalculationOptions calcOptions = new CalculationOptions
            {
                IgnoreError = true
            };

            // Recalculate all formulas in the workbook using the options
            workbook.CalculateFormula(calcOptions);

            // Output the evaluated value (will be 0 because the error is ignored)
            Console.WriteLine("A1 Value: " + sheet.Cells["A1"].Value);

            // Save the workbook to a file (optional)
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
