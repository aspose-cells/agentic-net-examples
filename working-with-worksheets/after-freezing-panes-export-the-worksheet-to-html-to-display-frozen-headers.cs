// Title: Freeze the first row of an Excel worksheet and export to HTML with a persistent header using Aspose.Cells for .NET
// AI Prompts: Use Aspose.Cells to freeze the top header row of a worksheet and then save the workbook as an HTML file that retains the frozen pane. | Generate HTML from an Excel file in C# where the header row stays fixed while scrolling, leveraging Worksheet.FreezePanes and SaveFormat.Html.
// Common Searches: C# Aspose.Cells freeze top row before exporting to HTML | How to keep header row static in HTML output from Excel using Aspose.Cells | Export Excel worksheet to HTML with frozen panes in .NET | Aspose.Cells preserve frozen panes when saving as HTML | Freeze panes in Excel and generate HTML with fixed header using C#
// Tags: Worksheet.FreezePanes for HTML rendering | HTML export with fixed header using Aspose.Cells | preserve frozen pane state in HTML output | Aspose.Cells SaveFormat.Html with frozen rows | C# generate HTML from Excel with locked header

using Aspose.Cells;

// Loads input.xlsx, freezes the first row of the first worksheet, and saves the workbook as output.html where the header remains fixed in the generated HTML.
class Program
{
    static void Main()
    {
        // Load the workbook (using the provided load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Freeze the first row (header) so it stays visible in HTML
        // Parameters: totalRows, totalColumns, rows, columns
        // Freeze 1 row, 0 columns; the scrollable area starts at cell A2 (row index 1)
        sheet.FreezePanes(1, 0, 1, 0);

        // Export the worksheet to HTML (using the provided save rule)
        // The frozen pane information is preserved in the generated HTML
        workbook.Save("output.html", SaveFormat.Html);
    }
}
