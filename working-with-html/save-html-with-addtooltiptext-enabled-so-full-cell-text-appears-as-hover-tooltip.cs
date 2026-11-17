// Title: Save an Excel workbook as HTML with full cell text displayed as a hover tooltip using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# enable AddTooltipText when exporting to HTML | save Excel worksheet as HTML with cell hover tooltip using Aspose.Cells | how to show full cell value as tooltip in HTML output from Aspose.Cells | HtmlSaveOptions AddTooltipText property example in .NET | force column truncation and display tooltip in Aspose.Cells HTML export
// Tags: Aspose.Cells HtmlSaveOptions AddTooltipText | export Excel to HTML with hover tooltip | C# set column width for tooltip display | generate HTML from workbook with full cell tooltip | Aspose.Cells tooltip for truncated cell content

using Aspose.Cells;
using System;

// The example creates a workbook, writes a long string to cell A1, narrows column A to cause truncation, enables the AddTooltipText flag in HtmlSaveOptions, and saves the file as output.html so the complete cell value appears as a hover tooltip in the generated HTML.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Put a long text into a cell to demonstrate tooltip
        sheet.Cells["A1"].PutValue("This is a very long text that exceeds the cell width and should appear as a tooltip when hovering over the cell in the HTML output.");

        // Set column width to force truncation in the HTML view
        sheet.Cells.SetColumnWidth(0, 15); // Column A width

        // Configure HTML save options to add tooltip text
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.AddTooltipText = true; // Enable full cell text as hover tooltip

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
