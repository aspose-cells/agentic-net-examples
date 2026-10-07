// Title: Export an Excel workbook to HTML with full cell text displayed as hover tooltip using Aspose.Cells for .NET
// AI Prompts: Generate an HTML file from a workbook and enable tooltips that show each cell’s complete value on mouse hover with Aspose.Cells. | Set HtmlSaveOptions.AddTooltipText to true in C# and save the workbook as HTML so truncated cells reveal their full text as a tooltip.
// Common Searches: Aspose.Cells C# enable tooltip for full cell text when saving to HTML | how to use HtmlSaveOptions AddTooltipText property in Aspose.Cells | export Excel to HTML with hover tooltips for long cell values Aspose.Cells | C# Aspose.Cells HTML export truncated column width tooltip example | display full cell content on mouse over in generated HTML using Aspose.Cells
// Tags: Aspose.Cells HTML export tooltip | HtmlSaveOptions AddTooltipText C# | hover tooltip for truncated cells | export Excel to HTML with full cell text | configure column width for tooltip display

using Aspose.Cells;

// The example creates a workbook, writes a long string into cell A1, narrows the column width to cause truncation, enables the AddTooltipText option in HtmlSaveOptions, and saves the file as output.html. The resulting HTML shows the full cell content as a hover tooltip for cells whose displayed text is clipped.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put a long text into a cell to demonstrate tooltip on overflow
        sheet.Cells["A1"].PutValue("This is a very long text that will not fit in the cell width and should appear as a tooltip when hovered.");

        // Set a small column width so the text is truncated in the view
        sheet.Cells.SetColumnWidth(0, 10); // width in characters

        // Configure HTML save options to include tooltip text for cells
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.AddTooltipText = true; // Enable full cell text as hover tooltip

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
