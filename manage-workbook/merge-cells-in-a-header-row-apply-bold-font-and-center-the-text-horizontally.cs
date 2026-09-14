// Title: How to merge header cells, set bold font, and center text horizontally in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that merges cells A1 through D1, applies a bold font style with horizontal center alignment, and saves the workbook as an .xlsx file. | Show how to create a Style object in Aspose.Cells, enable Font.IsBold, set HorizontalAlignment to Center, and apply the style to a merged range in a worksheet. | Provide a concise example that builds a new workbook, merges the first‑row cells, formats the header text, and outputs the file using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# merge first row cells and apply bold centered header style | How to set bold font and center alignment on a merged range with Aspose.Cells .NET | C# example for creating a merged header row in an Excel file using Aspose.Cells
// Tags: merge cells A1:D1 Aspose.Cells C# | bold text styling Aspose.Cells | horizontal center alignment Aspose.Cells | header row formatting Aspose.Cells .NET | export workbook to XLSX Aspose.Cells

using Aspose.Cells;
using System;

// Creates a new workbook, merges cells A1:D1, sets a bold font with horizontal center alignment on the merged range, and saves the file as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Put header text in the first cell
        cells["A1"].PutValue("Header Title");

        // Merge cells A1:D1 (first row, columns 1 to 4)
        cells.Merge(0, 0, 1, 4); // row index, column index, total rows, total columns

        // Create a style with bold font and centered horizontal alignment
        Style headerStyle = workbook.CreateStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.HorizontalAlignment = TextAlignmentType.Center;

        // Apply the style to the merged range (style applied to the first cell of the range)
        StyleFlag flag = new StyleFlag();
        flag.All = true;
        cells["A1"].SetStyle(headerStyle, flag);

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
