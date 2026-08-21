// Title: Aspose.Cells .NET – Verify formulas stay unchanged after default column deletion
// Description: Demonstrates how a SUM formula on a Summary sheet that references Data!A1:A3 remains identical after deleting column A from the Data sheet with the default DeleteColumn method, confirming that cross‑sheet references are not auto‑updated.
// Keywords: Aspose.Cells verify formula after column delete | default DeleteColumn behavior .NET | cross sheet formula unchanged | preserve formulas Aspose.Cells | column deletion without reference update
// Common Searches: Aspose.Cells keep formula after deleting column | Does DeleteColumn update references in other worksheets | Check formula integrity after column removal Aspose.Cells | Aspose.Cells default DeleteColumn does not adjust formulas
// Developer Intent: Confirm that a formula in another worksheet is not modified when a column is removed using the default DeleteColumn method.
// Use Cases: Automated validation that summary calculations remain accurate after source columns are programmatically removed. | Logging formula values before and after column deletion to detect unintended changes. | Ensuring data‑cleanup scripts do not break cross‑sheet references in financial models.
// AI Prompts: Write C# code with Aspose.Cells that deletes a column but leaves formulas in other sheets unchanged and logs the before/after formulas. | Show how to compare a cell's formula before and after calling DeleteColumn and output a boolean indicating if it changed. | Explain how to enable reference updating in Aspose.Cells when deleting columns, instead of preserving the original formulas.

using System;
using Aspose.Cells;

// Demonstrates how a SUM formula on a Summary sheet that references Data!A1:A3 remains identical after deleting column A from the Data sheet with the default DeleteColumn method, confirming that cross‑sheet references are not auto‑updated.
class Program
{
    static void Main()
    {
        // Create a new workbook with two worksheets
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";
        Worksheet summarySheet = workbook.Worksheets.Add("Summary");

        // Populate column A in the Data sheet
        dataSheet.Cells["A1"].PutValue(10);
        dataSheet.Cells["A2"].PutValue(20);
        dataSheet.Cells["A3"].PutValue(30);

        // Set a formula in the Summary sheet that references the Data sheet
        summarySheet.Cells["B1"].Formula = "=SUM(Data!A1:A3)";

        // Store the formula before column deletion
        string formulaBefore = summarySheet.Cells["B1"].Formula;
        Console.WriteLine("Formula before deletion: " + formulaBefore);

        // Delete column A (index 0) from the Data sheet using the default DeleteColumn method
        // This method does NOT update references in other worksheets
        dataSheet.Cells.DeleteColumn(0);

        // Retrieve the formula after deletion
        string formulaAfter = summarySheet.Cells["B1"].Formula;
        Console.WriteLine("Formula after deletion: " + formulaAfter);

        // Verify that the formula remained unchanged
        bool unchanged = formulaBefore == formulaAfter;
        Console.WriteLine("Formula unchanged: " + unchanged);

        // Save the workbook
        workbook.Save("FormulaCheck.xlsx");
    }
}
