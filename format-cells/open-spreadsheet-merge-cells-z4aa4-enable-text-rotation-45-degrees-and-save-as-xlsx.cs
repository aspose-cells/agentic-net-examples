// Title: Merge cells Z4:AA4 and set a 45-degree text rotation in an XLSX workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells merge Z4:AA4 cells and rotate text 45 degrees | How to apply 45° text orientation to a merged range with Aspose.Cells | Saving an Excel file after merging cells and setting text rotation using Aspose.Cells .NET | Aspose.Cells style merged cells with rotation angle in C# | Set text rotation on merged cells in an XLSX workbook using Aspose.Cells
// Tags: merge cells Z4 AA4 Aspose.Cells | apply text rotation Aspose.Cells C# | merged cell style Aspose.Cells | save workbook as xlsx Aspose.Cells | cell range formatting Aspose.Cells .NET

using Aspose.Cells;
using System;

// Loads input.xlsx, merges the Z4:AA4 range, applies a 45-degree text rotation to the merged cells, and saves the result as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells Z4:AA4
        // Row index for row 4 is 3 (zero‑based), column index for Z is 25, AA is 26
        // Merge(startRow, startColumn, totalRows, totalColumns)
        sheet.Cells.Merge(3, 25, 1, 2);

        // Set text rotation to 45 degrees for the merged range
        Style style = sheet.Cells[3, 25].GetStyle();
        style.RotationAngle = 45;
        sheet.Cells[3, 25].SetStyle(style); // merged cells share the style of the first cell

        // Save the modified workbook as XLSX
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
