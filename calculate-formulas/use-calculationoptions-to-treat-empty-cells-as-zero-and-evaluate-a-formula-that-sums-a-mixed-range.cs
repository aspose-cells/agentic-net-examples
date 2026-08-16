// Title: C# – Sum a range with empty cells as zero using Aspose.Cells CalculationOptions
// Description: Demonstrates how to calculate =SUM(A1:A5) in a workbook where some cells are blank. The example shows the default behavior (blank cells counted as zero) and how to pass a CalculationOptions object (e.g., toggling IgnoreError) without affecting the result.
// Keywords: Aspose.Cells CalculationOptions | C# SUM blank cells zero | Worksheet.CalculateFormula options | Aspose.Cells empty cell handling | Treat blank cells as zero Aspose.Cells
// Common Searches: Aspose.Cells treat blank cells as zero | C# calculate SUM with empty cells Aspose.Cells | CalculationOptions IgnoreError effect | Worksheet.CalculateFormula example C# | Sum range with missing values Aspose.Cells
// Developer Intent: Compute a SUM formula over a mixed range where empty cells are considered zero, using CalculationOptions to confirm that other settings do not change the outcome.
// Use Cases: Run the default calculation to verify that blanks are treated as zero and the sum equals 70. | Create a CalculationOptions instance, set IgnoreError, and pass it to CalculateFormula to show the sum remains unchanged. | Toggle the IgnoreError flag on the same options object and recalculate to confirm that this option does not affect empty‑cell handling.
// AI Prompts: Write C# code that uses Aspose.Cells CalculationOptions to ensure blank cells are counted as zero when evaluating any formula. | Modify the example to return the SUM result as a double instead of an object. | Explain why changing CalculationOptions.IgnoreError has no impact on the SUM of a range containing empty cells.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to calculate =SUM(A1:A5) in a workbook where some cells are blank. The example shows the default behavior (blank cells counted as zero) and how to pass a CalculationOptions object (e.g., toggling IgnoreError) without affecting the result.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate a mixed range: some cells have values, others are left empty
            sheet.Cells["A1"].PutValue(10);   // numeric
            sheet.Cells["A2"].PutValue(20);   // numeric
            // A3 is left empty
            sheet.Cells["A4"].PutValue(40);   // numeric
            // A5 is left empty

            // -----------------------------------------------------------------
            // 1) Default calculation (empty cells are ignored, treated as 0)
            // -----------------------------------------------------------------
            object defaultResult = sheet.CalculateFormula("=SUM(A1:A5)");
            Console.WriteLine("Default SUM(A1:A5) result: " + defaultResult); // Expected 70

            // -----------------------------------------------------------------
            // 2) Use CalculationOptions (demonstrating other available options)
            // -----------------------------------------------------------------
            CalculationOptions options = new CalculationOptions
            {
                // Example option: ignore errors during calculation
                IgnoreError = false
            };

            object optResult = sheet.CalculateFormula("=SUM(A1:A5)", options);
            Console.WriteLine("SUM(A1:A5) with CalculationOptions: " + optResult); // Expected 70

            // -----------------------------------------------------------------
            // 3) Demonstrate that changing another option does not affect the sum
            // -----------------------------------------------------------------
            options.IgnoreError = true; // toggle the option
            object optResult2 = sheet.CalculateFormula("=SUM(A1:A5)", options);
            Console.WriteLine("SUM(A1:A5) after toggling IgnoreError: " + optResult2); // Still 70

            // Save the workbook (optional, demonstrates lifecycle rule compliance)
            string outputPath = "SumWithEmptyCells.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
