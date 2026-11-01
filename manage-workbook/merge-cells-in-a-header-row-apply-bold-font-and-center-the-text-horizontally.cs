// Title: Merge header row cells, set bold font and center alignment with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge A1:D1 and apply bold centered header | C# code to merge cells and center text using Aspose.Cells | How to format a merged header row with bold and center alignment in Aspose.Cells for .NET | Set horizontal alignment and bold style on merged cells with Aspose.Cells C# | Create a merged header and style it in an Excel file using Aspose.Cells .NET
// Tags: merge cells header Aspose.Cells C# | apply bold style merged cell Aspose.Cells | center horizontal alignment Aspose.Cells | CellArea range merging Aspose.Cells | save workbook as xlsx Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, merges cells A1‑D1 on the first worksheet, writes "Report Header", applies a bold font, centers the text horizontally, and saves the file as MergedHeader.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Define the range to merge (e.g., cells A1 to D1)
        CellArea headerArea = new CellArea
        {
            StartRow = 0,   // Row index is zero‑based (0 = first row)
            StartColumn = 0, // Column index is zero‑based (0 = column A)
            EndRow = 0,
            EndColumn = 3   // D column
        };

        // Merge the defined range
        sheet.Cells.Merge(headerArea.StartRow, headerArea.StartColumn,
                          headerArea.EndRow - headerArea.StartRow + 1,
                          headerArea.EndColumn - headerArea.StartColumn + 1);

        // Set the header text
        Cell headerCell = sheet.Cells[headerArea.StartRow, headerArea.StartColumn];
        headerCell.PutValue("Report Header");

        // Apply bold font
        Style style = headerCell.GetStyle();
        style.Font.IsBold = true;

        // Center the text horizontally
        style.HorizontalAlignment = TextAlignmentType.Center;

        // Apply the style back to the cell
        headerCell.SetStyle(style);

        // Save the workbook to a file
        workbook.Save("MergedHeader.xlsx");
    }
}
