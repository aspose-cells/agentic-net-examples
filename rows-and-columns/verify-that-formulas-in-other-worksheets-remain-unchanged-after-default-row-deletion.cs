// Title: Verify cross‑sheet formula remains unchanged after default row deletion in Aspose.Cells for .NET
// Description: This C# example creates a workbook with two worksheets, adds numeric data to Sheet1, sets a SUM formula in Sheet2 that references Sheet1!A1:A3, calculates the workbook, deletes the second row of Sheet1 using Worksheet.Cells.DeleteRow (which does not adjust external references), recalculates, and compares the formula text before and after deletion to confirm it has not changed.
// Keywords: Aspose.Cells | C# | .NET | row deletion | Worksheet.Cells.DeleteRow | cross‑sheet formula | formula stability | reference unchanged | default DeleteRow behavior | Excel automation
// Common Searches: Aspose.Cells DeleteRow cross sheet reference | formula unchanged after row removal Aspose.Cells | does DeleteRow update external formulas .NET | verify formula text after deleting rows Aspose.Cells | C# Aspose.Cells keep formula reference
// Developer Intent: Ensure that a formula in another worksheet is not modified when a row is removed with the default DeleteRow method.
// Use Cases: Automated testing of workbook integrity after row deletions | Validating that external references stay static in financial models | Demonstrating the default behavior of DeleteRow for documentation | Creating a safeguard routine before applying bulk row deletions
// AI Prompts: Generate C# code using Aspose.Cells that deletes a row in Sheet1 and checks that a SUM formula in Sheet2 still points to the original range. | Write an NUnit test that asserts the formula string in a secondary worksheet does not change after Worksheet.Cells.DeleteRow is called on the source sheet. | Explain why Worksheet.Cells.DeleteRow does not adjust formulas in other worksheets and show how to programmatically verify formula stability.

using System;
using Aspose.Cells;

// This C# example creates a workbook with two worksheets, adds numeric data to Sheet1, sets a SUM formula in Sheet2 that references Sheet1!A1:A3, calculates the workbook, deletes the second row of Sheet1 using Worksheet.Cells.DeleteRow (which does not adjust external references), recalculates, and compares the formula text before and after deletion to confirm it has not changed.
class VerifyFormulaUnchangedAfterRowDeletion
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // First worksheet (default name is "Sheet1")
        Worksheet sheet1 = workbook.Worksheets[0];

        // Add a second worksheet (default name will be "Sheet2")
        Worksheet sheet2 = workbook.Worksheets[workbook.Worksheets.Add()];

        // Populate Sheet1 with some numeric data
        sheet1.Cells["A1"].PutValue(10);
        sheet1.Cells["A2"].PutValue(20);
        sheet1.Cells["A3"].PutValue(30);

        // In Sheet2 set a formula that references a range in Sheet1
        sheet2.Cells["B1"].Formula = "=SUM(Sheet1!A1:A3)";

        // Calculate formulas so that values are up‑to‑date
        workbook.CalculateFormula();

        // Store the original formula text for later comparison
        string originalFormula = sheet2.Cells["B1"].Formula;

        // Delete the second row (index 1) in Sheet1 using the default DeleteRow method.
        // This overload does NOT update references in other worksheets.
        sheet1.Cells.DeleteRow(1);

        // Re‑calculate formulas after the row deletion
        workbook.CalculateFormula();

        // Output the formula before and after deletion to verify it has not changed
        Console.WriteLine("Original formula: " + originalFormula);
        Console.WriteLine("Formula after deletion: " + sheet2.Cells["B1"].Formula);
        Console.WriteLine("Formula unchanged: " + (originalFormula == sheet2.Cells["B1"].Formula));

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("VerifyFormula.xlsx");
    }
}
