// Title: Aspose.Cells .NET: Verify formulas stay unchanged after default DeleteRow
// Description: Creates a workbook with a Data sheet and a Summary sheet, adds values and formulas that reference the Data sheet, calculates them, deletes the second row of the Data sheet using the default DeleteRow method, recalculates, and prints the formulas to show that external references remain unchanged.
// Keywords: Aspose.Cells | DeleteRow | C# | .NET | formula reference | external sheet formula | row deletion | calculate formulas | verify unchanged formulas | workbook manipulation
// Common Searches: Aspose.Cells DeleteRow keep formulas unchanged | C# verify formula strings after row deletion | external sheet references after DeleteRow Aspose.Cells | how to prevent formula updates when deleting rows in Aspose.Cells | Aspose.Cells default DeleteRow behavior
// Developer Intent: Confirm that using the default DeleteRow method does not modify formulas in other worksheets that reference the deleted rows.
// Use Cases: Automated test to ensure summary‑sheet calculations remain stable after cleaning source data rows. | Report generation where source rows are removed without affecting linked worksheet formulas. | Quality‑control script that compares formula strings before and after a DeleteRow operation.
// AI Prompts: Generate C# code with Aspose.Cells to delete multiple rows while preserving formulas that reference those rows from other worksheets. | Show how to programmatically compare formula strings before and after a DeleteRow call to verify they are unchanged. | Explain how to enable automatic reference updating for formulas when deleting rows in Aspose.Cells, if required.

using System;
using Aspose.Cells;

// Creates a workbook with a Data sheet and a Summary sheet, adds values and formulas that reference the Data sheet, calculates them, deletes the second row of the Data sheet using the default DeleteRow method, recalculates, and prints the formulas to show that external references remain unchanged.
class VerifyFormulaUnchanged
{
    static void Main()
    {
        // Create a new workbook with two worksheets
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";
        Worksheet summarySheet = workbook.Worksheets.Add("Summary");

        // Populate some values in the first worksheet
        dataSheet.Cells["A1"].PutValue(10);
        dataSheet.Cells["A2"].PutValue(20);
        dataSheet.Cells["A3"].PutValue(30);
        dataSheet.Cells["A4"].PutValue(40);

        // Add formulas in the second worksheet that reference the first worksheet
        summarySheet.Cells["A1"].Formula = "=Data!A1+5";
        summarySheet.Cells["A2"].Formula = "=SUM(Data!A1:A4)";
        summarySheet.Cells["A3"].Formula = "=Data!A3*2";

        // Calculate all formulas so that values are up‑to‑date
        workbook.CalculateFormula();

        // Display formulas before row deletion
        Console.WriteLine("Formulas before deletion:");
        Console.WriteLine($"A1: {summarySheet.Cells["A1"].Formula}");
        Console.WriteLine($"A2: {summarySheet.Cells["A2"].Formula}");
        Console.WriteLine($"A3: {summarySheet.Cells["A3"].Formula}");

        // Delete the second row (index 1) in the first worksheet using the default DeleteRow method
        // This does NOT update references in other worksheets
        dataSheet.Cells.DeleteRow(1);

        // Re‑calculate formulas after the deletion
        workbook.CalculateFormula();

        // Display formulas after row deletion – they should be unchanged
        Console.WriteLine("\nFormulas after deletion (should be unchanged):");
        Console.WriteLine($"A1: {summarySheet.Cells["A1"].Formula}");
        Console.WriteLine($"A2: {summarySheet.Cells["A2"].Formula}");
        Console.WriteLine($"A3: {summarySheet.Cells["A3"].Formula}");

        // Save the workbook (optional verification)
        workbook.Save("VerifyFormula.xlsx");
    }
}
