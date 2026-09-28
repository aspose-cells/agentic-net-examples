// Title: Set custom page margins and freeze the top worksheet row using Aspose.Cells for .NET
// AI Prompts: Write a C# snippet with Aspose.Cells that assigns specific left, right, top, and bottom margins (in inches) and then freezes the first row of the worksheet. | Generate .NET code that enables PrintHeadings, applies custom margins, and calls FreezePanes to keep the header row visible on screen and printed pages. | Create an Aspose.Cells workbook that configures printable area margins and freezes the top row before saving the file.
// Common Searches: Aspose.Cells C# set page margins in inches and freeze first row | how to keep header row visible while printing with Aspose.Cells .NET | freeze panes top row and set custom margins using Aspose.Cells API | print headings and custom margins in Excel workbook with Aspose.Cells | C# example for setting margins and freezing rows in Aspose.Cells
// Tags: custom page margins Aspose.Cells | freeze top row worksheet Aspose.Cells | enable PrintHeadings Aspose.Cells | set margins inches .NET | freeze panes C# Aspose.Cells

using System;
using Aspose.Cells;

// Creates a new workbook, applies 0.5‑0.75 inch margins, enables PrintHeadings, freezes the first row, and saves the file as CustomMarginAndFreeze.xlsx.
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

            // Set custom page margins (in inches)
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.LeftMargin = 0.5;    // Left margin
            pageSetup.RightMargin = 0.5;   // Right margin
            pageSetup.TopMargin = 0.75;    // Top margin
            pageSetup.BottomMargin = 0.75; // Bottom margin

            // Freeze the top row so it stays visible
            // FreezePanes(row, column, totalRows, totalColumns)
            // row/column define the first scrollable cell; totalRows/totalColumns define how many rows/columns to freeze
            sheet.FreezePanes(1, 0, 1, 0);

            // Ensure headings are printed (helps keep frozen rows on printed pages)
            pageSetup.PrintHeadings = true;

            // Save the workbook
            workbook.Save("CustomMarginAndFreeze.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
