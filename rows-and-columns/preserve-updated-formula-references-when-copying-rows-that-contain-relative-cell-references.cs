// Title: Copy rows and retain relative SUM formula references with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that copies a block of rows using Aspose.Cells and automatically updates any relative formulas in the copied rows. | Show how to confirm that a SUM formula shifts from A1:A5 to A6:A10 after invoking Cells.CopyRows in Aspose.Cells. | Provide an example that saves the workbook after copying rows while preserving formula behavior.
// Common Searches: Aspose.Cells copy rows keep relative formulas adjusted C# example | How does Cells.CopyRows handle formula references in .NET | C# copy Excel rows with SUM formula updating automatically using Aspose | Preserve relative cell references when duplicating rows in Aspose.Cells | CopyRows method formula range shift Aspose.Cells tutorial
// Tags: Aspose.Cells Cells.CopyRows formula adjustment | C# copy rows with relative formulas | Excel row duplication preserving SUM references | Aspose.Cells workbook save after row copy | relative cell reference handling in .NET Excel automation

using System;
using Aspose.Cells;

// The example creates a workbook, fills column A with values, adds a relative SUM formula in B1, copies the first five rows to start at row 6 using Cells.CopyRows (which updates the formula to reference A6:A10), prints both original and copied formulas, and saves the file as PreserveFormulaCopyRows.xlsx.
class PreserveFormulaReferences
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate column A with sample values (A1:A5)
        for (int i = 0; i < 5; i++)
        {
            cells[i, 0].PutValue(i + 1); // A1 = 1, A2 = 2, ...
        }

        // Set a relative formula in B1 that sums the values in column A
        cells["B1"].Formula = "=SUM(A1:A5)";

        // Copy the first five rows (0‑based index) to rows starting at index 5 (row 6)
        // This operation automatically updates relative references in the copied formulas
        cells.CopyRows(cells, 0, 5, 5);

        // Display the original and the copied formulas to verify reference adjustment
        Console.WriteLine("Original formula in B1: " + cells["B1"].Formula);
        Console.WriteLine("Copied formula in B6: " + cells["B6"].Formula);

        // Save the workbook
        workbook.Save("PreserveFormulaCopyRows.xlsx");
    }
}
