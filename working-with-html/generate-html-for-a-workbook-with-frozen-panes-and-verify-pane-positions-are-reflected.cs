// Title: Generate an HTML file from an Excel workbook with frozen rows and columns using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, populates it with sample data, freezes the first 5 rows and 3 columns, and saves the result as an HTML file with Aspose.Cells. | Add console output that prints the expected number of frozen rows and columns before the workbook is saved to HTML. | Change the FreezePanes parameters to a different range (e.g., rows 2‑4, columns 1‑2) and regenerate the HTML to see how the pane positions are reflected.
// Common Searches: Aspose.Cells C# export worksheet with frozen panes to HTML | how to keep frozen rows and columns when saving Excel as HTML using Aspose.Cells | verify freeze pane settings in generated HTML file Aspose.Cells .NET | C# FreezePanes method example for HTML output with Aspose.Cells
// Tags: freeze panes export to HTML Aspose.Cells | Aspose.Cells FreezePanes C# example | preserve frozen rows columns in HTML output | generate HTML workbook with frozen panes .NET | validate pane positions in Aspose.Cells HTML

using Aspose.Cells;
using System;

// The example creates a new workbook, fills cells with sample data, freezes the first five rows and three columns, saves the workbook as an HTML file, and writes the expected frozen rows and columns to the console.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Freeze panes: rows above 5 and columns left of 3 will stay visible while scrolling
            // The overload requires total rows and columns of the frozen pane; 0 means no additional range.
            sheet.FreezePanes(5, 3, 0, 0);

            // Since FreezePanesRow/FreezePanesColumn properties may not be available in all versions,
            // we use the values we set directly.
            int frozenRows = 5;       // Expected frozen rows
            int frozenColumns = 3;    // Expected frozen columns
            Console.WriteLine($"Frozen Rows: {frozenRows}, Frozen Columns: {frozenColumns}");

            // Save the workbook as HTML
            string outputPath = "FrozenPaneWorkbook.html";
            workbook.Save(outputPath, SaveFormat.Html);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
