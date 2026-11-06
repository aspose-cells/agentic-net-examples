// Title: Create an XLSX workbook, merge cells P2:P5, enable text wrap, and add an AutoFilter using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# merge cells P2 to P5 and wrap text | how to apply an AutoFilter to range A1:D10 with Aspose.Cells | save a workbook with merged cells as XLSX using Aspose.Cells .NET | set text wrapping on a merged range in Aspose.Cells C# | create a new workbook and define an AutoFilter with Aspose.Cells
// Tags: merge cells P2:P5 Aspose.Cells .NET | enable text wrap merged cell Aspose.Cells | define autofilter range A1:D10 Aspose.Cells | save workbook as XLSX Aspose.Cells | create workbook with merged cells Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, merges the range P2:P5 on the first worksheet, applies a style that enables text wrapping to the merged cell, sets an AutoFilter for the range A1:D10, and saves the file as MergedWrapAutoFilter.xlsx in XLSX format using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells P2:P5 (column P = index 15, rows 2-5 = indices 1-4)
        sheet.Cells.Merge(1, 15, 4, 1); // startRow, startColumn, totalRows, totalColumns

        // Create a style with wrap text enabled
        Style wrapStyle = workbook.CreateStyle();
        wrapStyle.IsTextWrapped = true;

        // Apply the wrap style to the merged cell (top-left cell of the merged range)
        sheet.Cells["P2"].SetStyle(wrapStyle);

        // Define an AutoFilter range (adjust as needed; here we use A1:D10)
        sheet.AutoFilter.Range = "A1:D10";

        // Save the workbook as XLSX
        workbook.Save("MergedWrapAutoFilter.xlsx", SaveFormat.Xlsx);
    }
}
