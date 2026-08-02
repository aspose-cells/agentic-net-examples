// Title: Aspose.Cells C# – Print Gridlines while Hiding Row and Column Headings
// Description: Demonstrates how to make gridlines visible on screen, enable their printing, and suppress both on‑screen and printed row/column headings for the active worksheet, then saves the workbook as Gridlines_NoHeadings.xlsx.
// Keywords: Aspose.Cells print gridlines | hide row column headers Aspose.Cells | disable headings printing | IsRowColumnHeadersVisible | PageSetup.PrintGridlines | C# Excel formatting | Aspose.Cells worksheet settings
// Common Searches: Aspose.Cells print gridlines without headings | C# hide row and column headers in Excel file | How to disable heading printing in Aspose.Cells | Show gridlines on print Aspose.Cells .NET | Aspose.Cells hide worksheet headers
// Developer Intent: Configure the current worksheet to print gridlines but omit row and column headings from both the view and the printed output.
// Use Cases: Create printable reports that show only data gridlines for a clean layout. | Generate Excel templates where users see no headers in the UI, yet gridlines appear on paper or PDF. | Prepare spreadsheets for PDF export with gridlines for visual separation while keeping the page header‑free.
// AI Prompts: Write C# code using Aspose.Cells to enable gridline printing and hide row/column headings in the active worksheet. | Provide a reusable method that toggles gridline visibility and heading printing based on Boolean parameters. | Show how to set IsRowColumnHeadersVisible, PrintGridlines, and PrintHeadings properties in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Demonstrates how to make gridlines visible on screen, enable their printing, and suppress both on‑screen and printed row/column headings for the active worksheet, then saves the workbook as Gridlines_NoHeadings.xlsx.
class EnableGridlinesDisableHeadings
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first (current) worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Make gridlines visible on screen
        worksheet.IsGridlinesVisible = true;

        // Enable printing of gridlines
        worksheet.PageSetup.PrintGridlines = true;

        // Disable printing of row and column headings
        worksheet.PageSetup.PrintHeadings = false;

        // Hide row and column headers in the worksheet view
        worksheet.IsRowColumnHeadersVisible = false;

        // Sample data (optional, to illustrate the settings)
        worksheet.Cells["A1"].PutValue("Sample Data");
        worksheet.Cells["B2"].PutValue(123);

        // Save the workbook
        workbook.Save("Gridlines_NoHeadings.xlsx");
    }
}
