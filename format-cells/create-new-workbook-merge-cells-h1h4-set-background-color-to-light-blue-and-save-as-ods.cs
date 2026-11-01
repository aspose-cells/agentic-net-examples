// Title: How to merge cells H1:H4, apply a light‑blue fill, and save as ODS using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# merge a column range and set background color | Save an Aspose.Cells workbook as ODS after formatting cells | Apply solid light blue fill to merged cells with Aspose.Cells .NET | Create a new workbook and merge cells H1 to H4 using Aspose.Cells | Set cell shading for a merged range in Aspose.Cells C#
// Tags: merge cells range Aspose.Cells C# | apply solid background color Aspose.Cells | save workbook as ODS Aspose.Cells | style flag cell shading Aspose.Cells | create range and apply style Aspose.Cells

using Aspose.Cells;
using System;
using System.Drawing;

// The example creates a new workbook, merges the range H1:H4 on the first worksheet, applies a solid light‑blue background to the merged cells, and saves the file in ODS format using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Merge cells H1:H4 (column H = index 7, rows 0-3)
            sheet.Cells.Merge(0, 7, 4, 1); // startRow, startColumn, totalRows, totalColumns

            // Create a style with light blue background
            Style style = workbook.CreateStyle();
            style.ForegroundColor = Color.LightBlue;
            style.Pattern = BackgroundType.Solid;

            // Apply the style to the merged range
            Aspose.Cells.Range range = sheet.Cells.CreateRange("H1:H4");
            StyleFlag flag = new StyleFlag();
            flag.CellShading = true; // apply background color
            range.ApplyStyle(style, flag);

            // Save the workbook as ODS format
            workbook.Save("output.ods", SaveFormat.Ods);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
