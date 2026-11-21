// Title: How to render solid continuous gridlines when exporting an Excel worksheet to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export worksheet to PDF with solid gridlines | make Excel gridlines appear as continuous lines in PDF using Aspose.Cells | C# set thin borders for all cells to get solid gridlines in PDF output | PrintGridlines true not showing solid lines Aspose.Cells PDF | apply style to used range for solid gridlines in PDF with Aspose.Cells
// Tags: solid gridlines PDF export Aspose.Cells | PrintGridlines property PDF rendering | thin cell border style Aspose.Cells | apply style to used range C# | continuous gridlines in PDF generation

using System;
using Aspose.Cells;

// The example creates a workbook, fills cells with data, enables PrintGridlines, applies a thin border style to the entire used range, and saves the sheet as a PDF so that the gridlines render as solid continuous lines.
class GridlinesPdfExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data to visualize gridlines
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Enable printing of gridlines
            sheet.PageSetup.PrintGridlines = true;

            // Apply a thin continuous border to all cells to make gridlines solid
            Style style = workbook.CreateStyle();
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            // Apply the style to the used range
            Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;
            usedRange.ApplyStyle(style, new StyleFlag { Borders = true });

            // Save the workbook as PDF; gridlines will appear as continuous lines
            workbook.Save("GridlinesOutput.pdf", SaveFormat.Pdf);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
