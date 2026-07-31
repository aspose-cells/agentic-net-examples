// Title: C# – Merge A1:D1 and Apply a SUM Formula with Aspose.Cells
// Description: Shows how to create a workbook, fill A2‑D2 with numbers, merge A1:D1 into a single header cell, assign the formula =SUM(A2:D2), recalculate the sheet, display the result, and save the file as MergedTotal.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# merge cells | Aspose.Cells set formula | SUM formula Aspose | Workbook.CalculateFormula | save Excel file Aspose | merged header total | Excel automation .NET
// Common Searches: Aspose.Cells merge cells and set formula | C# merge A1:D1 Aspose.Cells example | calculate sum in merged cell using Aspose | Workbook.CalculateFormula usage C# | save workbook with merged cells Aspose .NET
// Developer Intent: Combine a range into one cell and embed a SUM expression that totals a data row.
// Use Cases: Create a spanning header that automatically shows the total of a row of figures. | Build a financial summary where the top merged row reflects the sum of monthly values. | Design an invoice template with a merged cell that calculates the grand total of line items.
// AI Prompts: Generate C# code with Aspose.Cells to merge cells A1:D1, insert =SUM(A2:D2) in the merged cell, evaluate the formula, and save the workbook. | Explain how to use Workbook.CalculateFormula after merging cells and assigning a formula in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Shows how to create a workbook, fill A2‑D2 with numbers, merge A1:D1 into a single header cell, assign the formula =SUM(A2:D2), recalculate the sheet, display the result, and save the file as MergedTotal.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Fill some sample numbers that we will total
        cells["A2"].PutValue(10);
        cells["B2"].PutValue(20);
        cells["C2"].PutValue(30);
        cells["D2"].PutValue(40);

        // Merge cells A1:D1 (row 0, columns 0‑3)
        // Parameters: firstRow, firstColumn, totalRows (1‑based), totalColumns (1‑based)
        cells.Merge(0, 0, 1, 4);

        // Set a formula in the merged cell (address A1) to calculate the sum of A2:D2
        cells["A1"].Formula = "=SUM(A2:D2)";

        // Evaluate the formula so the result is stored in the cell
        workbook.CalculateFormula();

        // Output the calculated total to the console (optional)
        Console.WriteLine("Calculated total: " + cells["A1"].Value);

        // Save the workbook to a file
        workbook.Save("MergedTotal.xlsx");
    }
}
