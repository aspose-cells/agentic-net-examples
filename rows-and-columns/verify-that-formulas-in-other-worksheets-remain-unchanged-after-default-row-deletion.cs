// Title: Check that external sheet formulas stay unchanged after DeleteRow in Aspose.Cells for .NET
// Description: Creates a workbook with a Data sheet and a Summary sheet, fills Data!A1:A5 with numbers, sets a SUM formula on Summary referencing that range, records the formula text, deletes row 2 from Data using the default DeleteRow (which leaves external references intact), recalculates, and confirms the formula string is identical before and after the deletion.
// Keywords: Aspose.Cells | DeleteRow | formula unchanged | external reference | C# example | row deletion | worksheet formula stability | calculate formula | .NET
// Common Searches: Aspose.Cells DeleteRow keep formula reference | Does DeleteRow adjust formulas in other sheets | Verify formula text after row removal Aspose.Cells | C# check external sheet formula after DeleteRow | Aspose.Cells row deletion impact on formulas
// Developer Intent: Confirm that a formula on a different worksheet remains exactly the same after removing a row with the default DeleteRow method.
// Use Cases: Validate that external‑sheet formulas are not auto‑updated when rows are deleted. | Capture and compare a cell's formula string before and after a DeleteRow operation. | Demonstrate saving a workbook after confirming formula stability.
// AI Prompts: Provide C# code that records a cell's formula, deletes a row with DeleteRow, and verifies the formula text is unchanged using Aspose.Cells. | Show a snippet illustrating that the default DeleteRow method does not modify formulas in other worksheets and how to test this behavior. | Explain why DeleteRow leaves external references untouched and give a programmatic way to confirm it.

using System;
using Aspose.Cells;

// Creates a workbook with a Data sheet and a Summary sheet, fills Data!A1:A5 with numbers, sets a SUM formula on Summary referencing that range, records the formula text, deletes row 2 from Data using the default DeleteRow (which leaves external references intact), recalculates, and confirms the formula string is identical before and after the deletion.
class VerifyFormulaUnchanged
{
    static void Main()
    {
        // Create a new workbook with two worksheets
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";
        Worksheet summarySheet = workbook.Worksheets.Add("Summary");

        // Fill the first worksheet with sample numbers in column A (rows 1‑5)
        for (int i = 0; i < 5; i++)
        {
            dataSheet.Cells[i, 0].PutValue(i + 1); // A1..A5 = 1..5
        }

        // In the second worksheet set a formula that sums the range A1:A5 of the first sheet
        summarySheet.Cells["A1"].Formula = "=SUM(Data!A1:A5)";

        // Calculate formulas so the workbook has up‑to‑date values
        workbook.CalculateFormula();

        // Store the formula text before any deletion
        string formulaBefore = summarySheet.Cells["A1"].Formula;

        // Delete row 2 (zero‑based index 1) from the first worksheet using the default DeleteRow method
        // This overload does NOT update references in other worksheets
        dataSheet.Cells.DeleteRow(1);

        // Re‑calculate formulas (optional, does not affect the formula text)
        workbook.CalculateFormula();

        // Store the formula text after deletion
        string formulaAfter = summarySheet.Cells["A1"].Formula;

        // Output the results
        Console.WriteLine("Formula before deletion: " + formulaBefore);
        Console.WriteLine("Formula after deletion : " + formulaAfter);
        Console.WriteLine("Formula unchanged: " + (formulaBefore == formulaAfter));

        // Save the workbook (optional, just to demonstrate saving)
        workbook.Save("VerifyFormula.xlsx");
    }
}
