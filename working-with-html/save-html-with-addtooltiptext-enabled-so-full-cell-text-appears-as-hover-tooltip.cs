// Title: Export an Excel worksheet to HTML with full cell text displayed as a hover tooltip using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, inserts a long string into a cell, narrows the column to force truncation, enables tooltip text in HtmlSaveOptions, and saves the result as an HTML file. | Show how to configure Aspose.Cells HtmlSaveOptions so that each cell in the generated HTML includes a tooltip containing the complete cell value. | Explain the steps to make truncated Excel cells show their full content on mouse hover after exporting to HTML with Aspose.Cells.
// Common Searches: Aspose.Cells C# export to HTML with tooltip for long cell values | how to show full cell text on hover in HTML output from Aspose.Cells | enable tooltip option in HtmlSaveOptions to display complete cell content | C# set column width and add hover tooltip when saving workbook as HTML
// Tags: Aspose.Cells HtmlSaveOptions AddTooltipText | export Excel to HTML with cell tooltips .NET | display full cell content on hover Aspose.Cells | C# set column width for HTML rendering Aspose.Cells | generate tooltip for truncated cells Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Rendering;

// // Creates a workbook, writes a long string to cell A1, narrows column A to cause truncation, sets HtmlSaveOptions.AddTooltipText = true, and saves the worksheet as output.html so the complete text appears as a hover tooltip in the HTML view.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put a long text into a cell (will be truncated in HTML view)
        sheet.Cells["A1"].PutValue("This is a very long text that exceeds the cell width and should appear as a tooltip when exported to HTML.");

        // Set column width to force truncation in the HTML rendering
        sheet.Cells.SetColumnWidth(0, 10); // Column A

        // Configure HTML save options to add tooltip text
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.AddTooltipText = true; // Enable full cell text as hover tooltip

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
