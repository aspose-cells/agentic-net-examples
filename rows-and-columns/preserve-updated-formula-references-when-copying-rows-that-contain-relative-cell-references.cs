// Title: Copy Rows with Relative Formulas Preserved Using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, fill column A with numbers, assign a relative formula (=A2) to B1, copy the first row to a new position with Cells.CopyRow, and verify that the copied formula in B3 automatically updates to =A4. The example also calculates the formulas and saves the file.
// Keywords: Aspose.Cells | CopyRow | relative formula | preserve formula when copying rows | C# | .NET | Excel automation | adjust formula offsets | worksheet row duplication | formula reference shift
// Common Searches: Aspose.Cells copy row preserve formula | CopyRow method relative reference .NET | how to keep Excel formulas when duplicating rows with Aspose | C# example copy row with formulas | adjust formula offsets after row copy Aspose.Cells
// Developer Intent: Duplicate a worksheet row while ensuring that any relative formulas automatically adjust to reference the correct cells in the new location.
// Use Cases: Replicate a calculation row in a financial model so each copy references its own data row. | Programmatically copy template rows that contain formulas, preserving the offset to adjacent columns. | Generate repeated report sections where formulas must adapt to new row positions without manual editing.
// AI Prompts: Generate C# code using Aspose.Cells to copy a row that contains a relative formula and display the updated formula after copying. | Explain how Cells.CopyRow updates relative cell references and how to retrieve the new formula string. | Provide an example that copies multiple rows with formulas, recalculates the workbook, and prints the resulting values.

using Aspose.Cells;
using System;

// Demonstrates how to create a workbook, fill column A with numbers, assign a relative formula (=A2) to B1, copy the first row to a new position with Cells.CopyRow, and verify that the copied formula in B3 automatically updates to =A4. The example also calculates the formulas and saves the file.
class PreserveFormulaCopyRows
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate column A with sample numbers (A1:A10)
        for (int i = 0; i < 10; i++)
        {
            cells[i, 0].PutValue(i + 1);
        }

        // Set a relative formula in B1 that refers to the cell directly below in column A
        // When this row is copied, the reference should shift accordingly (e.g., B3 -> A4)
        cells["B1"].Formula = "=A2";

        // Copy the first row (index 0) to the third row (index 2)
        // The formula reference will be automatically adjusted to maintain its relative offset
        cells.CopyRow(cells, 0, 2);

        // Display the original and copied formulas to verify the adjustment
        Console.WriteLine("Original formula in B1: " + cells["B1"].Formula); // Expected: =A2
        Console.WriteLine("Copied formula in B3: " + cells["B3"].Formula);   // Expected: =A4

        // Calculate formulas to obtain resulting values
        workbook.CalculateFormula();
        Console.WriteLine("Value in B1: " + cells["B1"].Value); // Should be 2
        Console.WriteLine("Value in B3: " + cells["B3"].Value); // Should be 4

        // Save the workbook
        workbook.Save("PreserveFormulaCopyRows.xlsx");
    }
}
