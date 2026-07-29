// Title: Force a cell's formula result to zero without altering the formula – Aspose.Cells for .NET (C#)
// Description: Creates a workbook, assigns "=SUM(1,2,3)" to A1, then uses the SetFormula method with the same formula string and a custom value of 0 to keep the formula unchanged while overriding the calculated result. The sample prints the formula and the new value before saving the file.
// Keywords: Aspose.Cells | C# | SetFormula | force cell value zero | preserve formula | override calculated result | Excel formula manipulation | unit test workbook | programmatic Excel | cell value override
// Common Searches: Aspose.Cells keep formula unchanged set result to zero | SetFormula method override cell value without removing formula | C# force Excel cell result to zero while retaining formula | how to set custom value for a formula cell in Aspose.Cells | verify formula string after using SetFormula in .NET
// Developer Intent: The developer needs to retain a cell's original formula while programmatically setting its evaluated value to zero.
// Use Cases: Unit‑testing scenarios where the formula must stay intact but a known result is required. | Generating template workbooks that display placeholder zeros without breaking underlying calculations. | Creating reports that hide actual formula outcomes for presentation purposes while preserving the logic.
// AI Prompts: Write C# code with Aspose.Cells that assigns a formula to a cell and then forces its value to zero without changing the formula. | Explain how SetFormula works in Aspose.Cells and how to confirm the formula text remains unchanged after setting a custom result. | Show how to programmatically verify that a cell's formula string is unchanged after calling SetFormula with a new value.

using System;
using Aspose.Cells;

// Creates a workbook, assigns "=SUM(1,2,3)" to A1, then uses the SetFormula method with the same formula string and a custom value of 0 to keep the formula unchanged while overriding the calculated result. The sample prints the formula and the new value before saving the file.
class VerifyFormulaZero
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Access cell A1 and assign a formula
        Cell cell = worksheet.Cells["A1"];
        cell.Formula = "=SUM(1,2,3)"; // Original formula evaluates to 6

        // Keep the formula unchanged but set its calculated value to zero
        cell.SetFormula(cell.Formula, 0);

        // Verify that the formula string is still the same
        Console.WriteLine("Formula after SetFormula: " + cell.Formula);

        // Verify that the cell's value is now zero
        Console.WriteLine("Value after SetFormula: " + cell.Value);

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("VerifyFormulaZero.xlsx");
    }
}
