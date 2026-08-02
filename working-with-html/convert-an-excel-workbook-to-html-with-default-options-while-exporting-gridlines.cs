// Title: Convert Excel to HTML with Gridlines using Aspose.Cells for .NET (C#)
// Description: Shows how to enable worksheet gridlines, configure HtmlSaveOptions.ExportGridLines, and save a workbook as an HTML file with the library's default settings in C#.
// Keywords: Aspose.Cells | C# | Excel to HTML conversion | Export gridlines | HtmlSaveOptions | default HTML options | worksheet gridlines | save workbook as HTML
// Common Searches: Aspose.Cells export gridlines to HTML C# | How to save Excel as HTML with gridlines using .NET | HtmlSaveOptions ExportGridLines example | Convert workbook to HTML default settings Aspose | C# code to keep Excel gridlines in HTML output
// Developer Intent: Generate an HTML representation of an Excel workbook while preserving the visible gridlines, using Aspose.Cells for .NET with minimal configuration.
// Use Cases: Display spreadsheet data on a web page with the same layout as Excel, including gridlines for readability. | Create static HTML reports that can be emailed or archived without losing the grid structure. | Batch‑process multiple workbooks into HTML files for documentation portals while keeping a consistent appearance.
// AI Prompts: Modify the example to export gridlines for only selected worksheets. | Save the HTML output to a MemoryStream instead of a file while retaining gridlines. | Add custom CSS to the generated HTML to change fonts and colors while using HtmlSaveOptions.

using System;
using Aspose.Cells;

// Shows how to enable worksheet gridlines, configure HtmlSaveOptions.ExportGridLines, and save a workbook as an HTML file with the library's default settings in C#.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Make gridlines visible in the worksheet
        worksheet.IsGridlinesVisible = true;

        // Create HTML save options and enable gridline export
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportGridLines = worksheet.IsGridlinesVisible
        };

        // Save the workbook as HTML with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
