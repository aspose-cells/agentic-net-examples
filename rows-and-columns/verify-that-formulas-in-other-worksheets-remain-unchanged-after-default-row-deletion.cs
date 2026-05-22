using System;
using Aspose.Cells;

namespace VerifyFormulaUnchangedAfterRowDeletion
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook with two worksheets
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Add a second worksheet for formulas that reference the first sheet
            Worksheet summarySheet = workbook.Worksheets[workbook.Worksheets.Add()];
            summarySheet.Name = "Summary";

            // Populate the Data sheet with some values in column A (rows 1‑5)
            for (int i = 0; i < 5; i++)
            {
                dataSheet.Cells[i, 0].PutValue(i + 1); // A1=1, A2=2, ...
            }

            // Set a formula in the Summary sheet that references cells in the Data sheet
            // Example: =Data!A1+Data!A2
            summarySheet.Cells["A1"].Formula = "=Data!A1+Data!A2";

            // Capture the formula text before deletion
            string formulaBefore = summarySheet.Cells["A1"].Formula;
            Console.WriteLine("Formula before deletion: " + formulaBefore);

            // Delete the first row (index 0) in the Data sheet using the default DeleteRow method
            // This method does NOT update references in other worksheets
            dataSheet.Cells.DeleteRow(0);

            // Capture the formula text after deletion
            string formulaAfter = summarySheet.Cells["A1"].Formula;
            Console.WriteLine("Formula after deletion:  " + formulaAfter);

            // Verify that the formula remained unchanged
            if (formulaBefore == formulaAfter)
            {
                Console.WriteLine("Success: Formula in other worksheet remained unchanged.");
            }
            else
            {
                Console.WriteLine("Failure: Formula was altered.");
            }

            // Save the workbook (optional, just to complete the lifecycle)
            workbook.Save("FormulaVerificationResult.xlsx");
        }
    }
}