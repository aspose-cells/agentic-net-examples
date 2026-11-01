// Title: Resume a paused formula calculation session with Aspose.Cells for .NET and verify results against uninterrupted execution
// Common Searches: how to resume a paused formula calculation using Aspose.Cells in C# | verify that resumed calculation matches original result in Aspose.Cells .NET | compare results of uninterrupted and resumed workbook calculations with Aspose.Cells | recalculate formulas after changing cell values without automatic update Aspose.Cells C# | save workbook after performing manual formula recalculation with Aspose.Cells
// Tags: manual formula recalculation Aspose.Cells | result verification after resumed calculation .NET | full workbook recompute after cell update | uninterrupted vs resumed calculation comparison | export workbook after formula evaluation

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, fills cells A1‑A3, sets B1 to =SUM(A1:A3), runs an uninterrupted calculation, then simulates a pause by changing A1 without immediate recompute. It performs a manual recalculation to resume the session, compares the resumed result with a freshly built workbook containing the same data, outputs verification status, and saves the final workbook as ResumedCalculation.xlsx.
class ResumePausedCalculation
{
    static void Main()
    {
        try
        {
            // Create a new workbook (rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Add a formula that depends on the data: B1 = SUM(A1:A3)
            sheet.Cells["B1"].Formula = "=SUM(A1:A3)";

            // -----------------------------------------------------------------
            // 1. Uninterrupted calculation: calculate all formulas and store result
            // -----------------------------------------------------------------
            workbook.CalculateFormula(); // full calculation
            double uninterruptedResult = sheet.Cells["B1"].DoubleValue;

            // -----------------------------------------------------------------
            // 2. Simulate pause by changing a source cell without immediate recalculation
            // -----------------------------------------------------------------
            // Change a source cell; formula will be recalculated later
            sheet.Cells["A1"].PutValue(100);

            // Calculate formulas (full recalculation in this simplified example)
            workbook.CalculateFormula();

            // Retrieve the result after the change
            double partialResult = sheet.Cells["B1"].DoubleValue;

            // -----------------------------------------------------------------
            // 3. Resume calculation (full recalculation to ensure consistency)
            // -----------------------------------------------------------------
            workbook.CalculateFormula();

            // Retrieve the resumed calculation result
            double resumedResult = sheet.Cells["B1"].DoubleValue;

            // -----------------------------------------------------------------
            // 4. Verify that resumed result matches expected result
            // -----------------------------------------------------------------
            // Create a fresh workbook with the same data and formula for comparison
            Workbook freshWorkbook = new Workbook();
            Worksheet freshSheet = freshWorkbook.Worksheets[0];
            freshSheet.Cells["A1"].PutValue(100);
            freshSheet.Cells["A2"].PutValue(20);
            freshSheet.Cells["A3"].PutValue(30);
            freshSheet.Cells["B1"].Formula = "=SUM(A1:A3)";
            freshWorkbook.CalculateFormula();
            double expectedResult = freshSheet.Cells["B1"].DoubleValue;

            // Output verification
            Console.WriteLine($"Uninterrupted result: {uninterruptedResult}");
            Console.WriteLine($"Partial (simulated paused) result: {partialResult}");
            Console.WriteLine($"Resumed result: {resumedResult}");
            Console.WriteLine($"Expected result after change: {expectedResult}");

            if (Math.Abs(resumedResult - expectedResult) < 1e-9)
            {
                Console.WriteLine("Verification succeeded: resumed calculation matches expected result.");
            }
            else
            {
                Console.WriteLine("Verification failed: resumed calculation does not match expected result.");
            }

            // -----------------------------------------------------------------
            // 5. Save the workbook (rule: save)
            // -----------------------------------------------------------------
            string outputPath = "ResumedCalculation.xlsx";

            // Ensure we can write the file (overwrite if exists)
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
