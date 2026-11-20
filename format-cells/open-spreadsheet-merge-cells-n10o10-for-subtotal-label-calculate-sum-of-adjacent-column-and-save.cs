// Title: Merge cells N10:O10, insert a Subtotal label, and compute column M total with Aspose.Cells for .NET
// Common Searches: how to merge a range of cells and add a label using Aspose.Cells in C# | Aspose.Cells C# set SUM formula for column M and place result in P10 | C# code to merge N10 and O10 then calculate subtotal with Aspose.Cells | using Aspose.Cells to programmatically add a subtotal row in an Excel worksheet | Aspose.Cells merge cells and apply formula without opening Excel
// Tags: merge cells Aspose.Cells C# | set cell formula Aspose.Cells | subtotal label merged range Aspose.Cells | calculate column sum Aspose.Cells | save workbook after modifications Aspose.Cells

using Aspose.Cells;
using System;

// Loads input.xlsx, merges cells N10:O10, writes "Subtotal" into the merged cells, sets a SUM(M1:M9) formula in P10, and saves the workbook as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells N10:O10 (row index 9, column index 13, 1 row, 2 columns)
        sheet.Cells.Merge(9, 13, 1, 2);
        // Set the subtotal label in the merged cells
        sheet.Cells["N10"].PutValue("Subtotal");

        // Calculate the sum of the adjacent column (column M) for rows 1 through 9
        // Place the result in the cell right after the merged range (P10)
        sheet.Cells["P10"].Formula = "SUM(M1:M9)";

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
