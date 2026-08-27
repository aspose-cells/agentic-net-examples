// Title: Delete a column in one worksheet while keeping cross‑sheet formulas unchanged using Aspose.Cells for .NET
// AI Prompts: Use Worksheet.Cells.DeleteColumn to remove the first column of Sheet1 and print the formula in Sheet2!C1 before and after the deletion. | Demonstrate how to prevent external references from being updated when a column is deleted in Aspose.Cells and verify the referenced formula stays the same. | Recalculate the workbook after the column removal and output the resulting value of the unchanged cross‑sheet formula.
// Common Searches: Aspose.Cells .NET delete column without changing formulas that reference it from another worksheet | preserve external worksheet formula references after column deletion using Aspose.Cells | verify that a formula in Sheet2 still points to original cells after deleting column A in Sheet1 with Aspose.Cells | C# example of Worksheet.Cells.DeleteColumn not updating cross‑sheet references
// Tags: Worksheet.Cells.DeleteColumn without updating external references | cross‑worksheet formula preservation Aspose.Cells | verify formula unchanged after column deletion C# | Aspose.Cells workbook recalculate after column shift | Excel column removal impact on formulas Aspose.Cells

using System;
using Aspose.Cells;

namespace VerifyFormulaUnchangedAfterColumnDeletion
{
    // Shows how to delete the first column of Sheet1 using DeleteColumn (which does not adjust external references), then confirms that a formula in Sheet2 referencing Sheet1 remains unchanged before and after the deletion, recalculates the workbook, and saves the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with two worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];               // First worksheet
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");    // Second worksheet

            // Populate Sheet1 with sample data in columns A and B
            sheet1.Cells["A1"].PutValue(10);
            sheet1.Cells["B1"].PutValue(20);
            sheet1.Cells["A2"].PutValue(30);
            sheet1.Cells["B2"].PutValue(40);

            // In Sheet2, add a formula that references cells in Sheet1
            // Formula: =Sheet1!A1 + Sheet1!B1
            sheet2.Cells["C1"].Formula = "=Sheet1!A1+Sheet1!B1";

            // Display the formula before column deletion
            Console.WriteLine("Formula in Sheet2!C1 before deletion: " + sheet2.Cells["C1"].Formula);

            // Delete the first column (A) in Sheet1 without updating references in other worksheets
            // Using DeleteColumn(int) which does NOT update external references
            sheet1.Cells.DeleteColumn(0);

            // Display the formula after column deletion to verify it is unchanged
            Console.WriteLine("Formula in Sheet2!C1 after deletion: " + sheet2.Cells["C1"].Formula);

            // Optionally, recalculate formulas to see the effect of the deleted column
            workbook.CalculateFormula();

            // Show the calculated value after deletion (should reflect the new reference if it were updated,
            // but because we did not update references, the formula still points to the original cells,
            // which now correspond to the original B column shifted to A)
            Console.WriteLine("Calculated value in Sheet2!C1 after deletion: " + sheet2.Cells["C1"].StringValue);

            // Save the workbook
            workbook.Save("VerifyFormulaUnchanged.xlsx");
        }
    }
}
