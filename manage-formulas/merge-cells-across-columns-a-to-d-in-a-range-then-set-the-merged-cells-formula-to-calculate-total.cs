// Title: Merging cells A1:D1 and setting a SUM formula using Aspose.Cells for .NET in C#
// Common Searches: Aspose.Cells C# merge a range of cells and add a SUM formula | C# code to merge A1:D1 in Excel and calculate total of A2:D2 with Aspose.Cells | How to set a formula in a merged cell using Aspose.Cells for .NET | Create a merged header row with total calculation in Aspose.Cells C# example | Saving workbook after merging cells and applying a formula with Aspose.Cells
// Tags: merge cells range Aspose.Cells C# | assign SUM formula to merged cell Aspose.Cells | populate sample values A2:D2 Aspose.Cells | save workbook as XLSX using Aspose.Cells .NET | Excel header row merging with formula Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, merges cells A1 through D1 on the first worksheet, assigns a SUM formula that totals the values in A2:D2, inserts sample numeric values into those cells, and saves the file as MergedCellWithFormula.xlsx.
class MergeCellsAndSetFormula
{
    static void Main()
    {
        // Create a new workbook.
        Workbook workbook = new Workbook();

        // Access the first worksheet.
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells across columns A to D in the first row (A1:D1).
        // Parameters: firstRow, firstColumn, totalRows, totalColumns
        worksheet.Cells.Merge(0, 0, 1, 4); // Merges A1:D1

        // Set a formula in the merged cell to calculate the total of the cells below (A2:D2).
        worksheet.Cells["A1"].Formula = "=SUM(A2:D2)";

        // Optionally, put some sample numbers in A2:D2 for demonstration.
        worksheet.Cells["A2"].PutValue(10);
        worksheet.Cells["B2"].PutValue(20);
        worksheet.Cells["C2"].PutValue(30);
        worksheet.Cells["D2"].PutValue(40);

        // Save the workbook to a file.
        workbook.Save("MergedCellWithFormula.xlsx");
    }
}
