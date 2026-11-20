// Title: Resume a paused formula calculation in Aspose.Cells for .NET and verify that the result matches an uninterrupted calculation
// Common Searches: Aspose.Cells C# resume formula calculation after modifying cells | compare results of paused and continuous workbook calculation using Aspose.Cells | verify that CalculateFormula returns same value after clearing and recalculating a cell in .NET | recalculate workbook after cell update and check consistency with Aspose.Cells | how to force recalculation of a SUM formula after changing input values in Aspose.Cells
// Tags: resume paused calculation Aspose.Cells | recalculate workbook after cell update C# | validate formula consistency Aspose.Cells | clear cell value to force recalculation Aspose.Cells | compare uninterrupted and resumed calculation results

using System;
using Aspose.Cells;

// The example creates a workbook, fills A1‑A3 with numbers, assigns a SUM formula to B1, and calculates it. It stores the result, clears B1, changes A1, recalculates, and then verifies that the resumed calculation produces the same value as the uninterrupted run before saving the file.
class ResumePausedCalculation
{
    static void Main()
    {
        try
        {
            // ---------- Create a workbook (create rule) ----------
            Workbook workbook = new Workbook(); // creates a new workbook

            // Add sample data and formulas
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some numbers
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // Add a formula that sums the range
            cells["B1"].Formula = "=SUM(A1:A3)";

            // ---------- Uninterrupted calculation ----------
            // Perform full calculation in one go
            workbook.CalculateFormula();

            // Store the result for later comparison
            double uninterruptedResult = cells["B1"].DoubleValue;

            // ---------- Simulate pause ----------
            // Clear the result to force recalculation later
            cells["B1"].PutValue(null);

            // Simulate other operations while "paused"
            cells["A1"].PutValue(15);

            // ---------- Resume calculation ----------
            // Explicitly recalculate after modifications
            workbook.CalculateFormula();

            // Retrieve the result after paused/resumed calculation
            double resumedResult = cells["B1"].DoubleValue;

            // ---------- Verify that both results match ----------
            if (Math.Abs(uninterruptedResult - resumedResult) < 1e-9)
            {
                Console.WriteLine("Success: Results match. Value = " + resumedResult);
            }
            else
            {
                Console.WriteLine("Failure: Results differ.");
                Console.WriteLine("Uninterrupted: " + uninterruptedResult);
                Console.WriteLine("Resumed: " + resumedResult);
            }

            // ---------- Save the workbook (save rule) ----------
            string outputPath = "Result.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine("Workbook saved to " + outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
