// Title: Export Workbook with Frozen Panes to HTML and Verify Pane Positions – Aspose.Cells for .NET
// Description: Creates a new workbook, populates sample data, freezes panes at cell C3 (2 rows × 2 columns), reads the frozen pane settings with GetFreezedPanes, prints the results, and saves the worksheet as an HTML file that retains the frozen rows and columns.
// Keywords: Aspose.Cells freeze panes HTML | C# export Excel to HTML | GetFreezedPanes example | freeze panes verification | Aspose.Cells SaveFormat.Html | worksheet freeze panes C#
// Common Searches: how to freeze panes and export to html using aspose.cells | get freezed panes coordinates c# aspose.cells | save excel with frozen rows and columns as html | aspocells verify frozen pane settings | c# generate html preview of excel with frozen panes
// Developer Intent: Create a workbook, apply freeze panes, confirm their coordinates, and export the sheet to HTML while preserving the frozen view.
// Use Cases: Publish financial reports online where header rows and columns stay visible in the HTML view. | Automated testing to ensure frozen pane configuration matches design specifications before releasing HTML output. | Generate interactive HTML dashboards from Excel files that keep key rows/columns fixed for easier navigation.
// AI Prompts: Write C# code with Aspose.Cells to freeze panes at D5, retrieve the pane details, and save the worksheet as HTML. | Explain the parameters returned by GetFreezedPanes in Aspose.Cells and how to interpret them. | Provide suggestions for styling the HTML output to highlight frozen rows and columns.

using System;
using Aspose.Cells;

namespace AsposeCellsFreezePaneHtmlDemo
{
    // Creates a new workbook, populates sample data, freezes panes at cell C3 (2 rows × 2 columns), reads the frozen pane settings with GetFreezedPanes, prints the results, and saves the worksheet as an HTML file that retains the frozen rows and columns.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Freeze panes at cell C3 (row index 2, column index 2) with 2 frozen rows and 2 frozen columns
            sheet.FreezePanes("C3", 2, 2);

            // Verify the freeze pane settings
            int frozenRow, frozenColumn, frozenRows, frozenColumns;
            bool hasFreeze = sheet.GetFreezedPanes(out frozenRow, out frozenColumn, out frozenRows, out frozenColumns);

            Console.WriteLine($"Has frozen panes: {hasFreeze}");
            if (hasFreeze)
            {
                Console.WriteLine($"Freeze position - Row: {frozenRow}, Column: {frozenColumn}");
                Console.WriteLine($"Frozen rows: {frozenRows}, Frozen columns: {frozenColumns}");
            }

            // Save the workbook as HTML to visualize the frozen panes in a browser
            workbook.Save("FreezePaneDemo.html", SaveFormat.Html);

            Console.WriteLine("Workbook saved as HTML. Open 'FreezePaneDemo.html' to view the frozen panes.");
        }
    }
}
