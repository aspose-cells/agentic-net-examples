// Title: C# – Apply a custom style (font, fill, borders) to a UnionRange of two non‑adjacent ranges using Aspose.Cells
// Description: Demonstrates how to create a workbook, define two separate ranges (A1:B2 and C3:D4), merge them with UnionRanges, build a Style with Calibri bold white font, dark‑blue solid fill, and thin yellow borders, and apply that style to the combined range before saving the file.
// Keywords: Aspose.Cells | C# | UnionRange | custom cell style | font formatting | cell fill | border styling | non‑adjacent ranges | Excel automation .NET | Style object | range union example
// Common Searches: Aspose.Cells apply style to multiple ranges | How to style a UnionRange in C# | Set borders for non‑contiguous cells Aspose.Cells | Create custom cell style Aspose.Cells .NET | UnionRanges method example
// Developer Intent: Create a single visual style and apply it to the area formed by two distinct cell ranges.
// Use Cases: Highlight related sections of a report that are located in separate blocks of a worksheet. | Give consistent header formatting to multiple tables that are not contiguous. | Visually group cells across different areas for easier navigation or printing.
// AI Prompts: Generate C# code that unions the ranges A1:B2 and C3:D4 and applies a style with red borders and light gray fill using Aspose.Cells. | Show how to create a custom style with italic font and apply it to a UnionRange consisting of three separate ranges in a workbook. | Explain how to change the border color of an existing UnionRange style after it has been applied in Aspose.Cells.

using Aspose.Cells;
using System;
using System.Drawing;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// Demonstrates how to create a workbook, define two separate ranges (A1:B2 and C3:D4), merge them with UnionRanges, build a Style with Calibri bold white font, dark‑blue solid fill, and thin yellow borders, and apply that style to the combined range before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data in the two ranges
            sheet.Cells["A1"].PutValue("First");
            sheet.Cells["B2"].PutValue(123);
            sheet.Cells["C3"].PutValue("Second");
            sheet.Cells["D4"].PutValue(456);

            // Define the two ranges to be united
            AsposeRange range1 = sheet.Cells.CreateRange("A1:B2");
            AsposeRange range2 = sheet.Cells.CreateRange("C3:D4");

            // Create a UnionRange that combines the two ranges
            UnionRange union = range1.UnionRanges(new AsposeRange[] { range2 });

            // Create a custom style with font, fill, and borders
            Style style = workbook.CreateStyle();

            // Font settings
            style.Font.Name = "Calibri";
            style.Font.Size = 12;
            style.Font.IsBold = true;
            style.Font.Color = Color.White;

            // Fill settings
            style.ForegroundColor = Color.DarkBlue;
            style.Pattern = BackgroundType.Solid;

            // Border settings (top, bottom, left, right)
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.TopBorder].Color = Color.Yellow;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].Color = Color.Yellow;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].Color = Color.Yellow;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].Color = Color.Yellow;

            // Apply the custom style to the union range
            union.SetStyle(style);

            // Save the workbook to visualize the result
            workbook.Save("UnionRangeCustomStyle.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
