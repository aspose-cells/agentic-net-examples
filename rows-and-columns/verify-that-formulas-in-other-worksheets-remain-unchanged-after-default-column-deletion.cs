// Title: Check cross‑sheet formula stability after deleting a column with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add a formula in Sheet2 that references Sheet1, delete column A in Sheet1 using DeleteColumn (without updateReference), and verify that the formula text in Sheet2 stays unchanged.
// Keywords: Aspose.Cells | C# | DeleteColumn | cross‑sheet formula | formula integrity | column deletion | reference preservation | workbook calculation
// Common Searches: Aspose.Cells delete column without updating formulas | verify formula unchanged after column removal C# | cross sheet formula reference after structural change | how to keep formulas stable when deleting columns in Aspose.Cells
// Developer Intent: Confirm that removing a column in one worksheet does not alter formulas that reference that worksheet from other sheets.
// Use Cases: Unit test to ensure cross‑sheet formulas remain intact after column deletions. | Pre‑processing workbooks where columns are stripped but external calculations must stay accurate. | Generating validation reports that compare formula strings before and after structural modifications.
// AI Prompts: Write C# code using Aspose.Cells that deletes column A in Sheet1 without updating references and asserts that a formula in Sheet2 referencing Sheet1 stays unchanged. | Show how to capture a formula's text before and after deleting a column in Aspose.Cells and compare the values. | Explain why DeleteColumn without the updateReference parameter leaves external formulas untouched in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaVerification
{
    // Shows how to create a workbook, add a formula in Sheet2 that references Sheet1, delete column A in Sheet1 using DeleteColumn (without updateReference), and verify that the formula text in Sheet2 stays unchanged.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (Sheet1) and add a second worksheet (Sheet2)
            Worksheet sheet1 = workbook.Worksheets[0];
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Populate Sheet1 with sample data in columns A and B
            sheet1.Cells["A1"].PutValue(10);
            sheet1.Cells["B1"].PutValue(20);

            // In Sheet2, set a formula that references cells from Sheet1
            sheet2.Cells["A1"].Formula = "=Sheet1!A1+Sheet1!B1";

            // Calculate formulas so that any dependent values are up‑to‑date
            workbook.CalculateFormula();

            // Capture the formula text before deleting any column
            string formulaBefore = sheet2.Cells["A1"].Formula;

            // Delete the first column (index 0) in Sheet1 without updating references
            // This uses the DeleteColumn method that does NOT have the updateReference parameter,
            // meaning references in other worksheets remain unchanged.
            sheet1.Cells.DeleteColumn(0);

            // Capture the formula text after the column deletion
            string formulaAfter = sheet2.Cells["A1"].Formula;

            // Output the results to verify that the formula has not changed
            Console.WriteLine("Formula before deletion: " + formulaBefore);
            Console.WriteLine("Formula after deletion:  " + formulaAfter);
            Console.WriteLine("Formula unchanged: " + (formulaBefore == formulaAfter));

            // Save the workbook (optional, just to inspect the file if needed)
            workbook.Save("FormulaCheckAfterColumnDeletion.xlsx");
        }
    }
}
