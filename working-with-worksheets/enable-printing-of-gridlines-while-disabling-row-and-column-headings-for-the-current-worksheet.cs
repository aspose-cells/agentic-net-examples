using System;
using Aspose.Cells;

class EnableGridlinesDisableHeadings
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet (current worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure gridlines are visible on the sheet (optional for UI)
        worksheet.IsGridlinesVisible = true;

        // Enable printing of gridlines
        worksheet.PageSetup.PrintGridlines = true;

        // Disable printing of row and column headings
        worksheet.PageSetup.PrintHeadings = false;

        // Also hide row/column headers in the worksheet view
        worksheet.IsRowColumnHeadersVisible = false;

        // Save the workbook
        workbook.Save("GridlinesEnabled_HeadingsDisabled.xlsx");
    }
}