// Title: Aspose.Cells for .NET – Export Excel to HTML with Fixed Pixel Column Widths (WidthScalable = false)
// Description: This example creates a workbook, assigns column A (150 px) and column B (80 px) using SetColumnWidthPixel, disables automatic scaling via HtmlSaveOptions.WidthScalable = false, and saves the sheet as FixedWidthOutput.html where column sizes are rendered as exact pixel values.
// Keywords: Aspose.Cells HTML export | SetColumnWidthPixel | WidthScalable false | fixed column width HTML | C# Aspose.Cells example | pixel column width
// Common Searches: Aspose.Cells export HTML fixed column width | HtmlSaveOptions WidthScalable property C# | Set column width in pixels before HTML save | Disable column width scaling Aspose.Cells
// Developer Intent: Produce an HTML representation of an Excel worksheet where each column retains a predefined pixel width.
// Use Cases: Web dashboards that require a stable column layout across browsers. | Email templates with HTML tables that must keep exact column dimensions. | Printable HTML reports where column spacing cannot be altered by viewport size.
// AI Prompts: How do I set column widths in pixels and turn off width scaling when saving a workbook to HTML with Aspose.Cells for .NET? | Give me a C# code sample that verifies the generated HTML contains the expected pixel values after using WidthScalable = false. | Explain the effect of HtmlSaveOptions.WidthScalable on column rendering and how to enforce fixed pixel widths.

using System;
using Aspose.Cells;

namespace AsposeCellsWidthScalableDemo
{
    // This example creates a workbook, assigns column A (150 px) and column B (80 px) using SetColumnWidthPixel, disables automatic scaling via HtmlSaveOptions.WidthScalable = false, and saves the sheet as FixedWidthOutput.html where column sizes are rendered as exact pixel values.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set some sample data
            cells["A1"].PutValue("Fixed Width Column");
            cells["B1"].PutValue("Another Column");

            // Set column widths in pixels (fixed values)
            // Column A (index 0) to 150 pixels, Column B (index 1) to 80 pixels
            cells.SetColumnWidthPixel(0, 150);
            cells.SetColumnWidthPixel(1, 80);

            // Configure HTML save options to disable width scaling
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.WidthScalable = false; // Ensure column widths are exported as fixed pixel values

            // Save the workbook as HTML with the specified options
            workbook.Save("FixedWidthOutput.html", htmlOptions);
        }
    }
}
