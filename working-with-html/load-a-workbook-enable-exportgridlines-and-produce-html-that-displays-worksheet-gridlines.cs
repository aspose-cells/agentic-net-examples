// Title: Generate HTML from an Excel workbook with visible gridlines using Aspose.Cells for .NET
// AI Prompts: Create C# code that loads an .xlsx file, sets HtmlSaveOptions.ExportGridLines to true, and saves the workbook as an HTML file with gridlines displayed. | Demonstrate how to configure Aspose.Cells HtmlSaveOptions to preserve Excel worksheet gridlines when exporting to HTML in a .NET application.
// Common Searches: Aspose.Cells export Excel to HTML with gridlines C# example | How to enable gridlines in HTML output using Aspose.Cells HtmlSaveOptions | C# code to save workbook as HTML showing Excel cell borders with Aspose.Cells | Export worksheet gridlines to HTML using Aspose.Cells for .NET
// Tags: Aspose.Cells HtmlSaveOptions gridlines export | C# export Excel workbook to HTML with gridlines | Excel worksheet gridlines in HTML output using Aspose.Cells | Preserve cell borders when converting Excel to HTML in .NET

using System;
using Aspose.Cells;

// The program loads an existing Excel workbook, configures HtmlSaveOptions.ExportGridLines to true to include worksheet gridlines, and saves the workbook as an HTML file that displays the original Excel gridlines.
class ExportGridLinesToHtml
{
    static void Main()
    {
        // Load an existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to export grid lines
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportGridLines = true // Enable grid lines in the generated HTML
        };

        // Save the workbook as an HTML file with grid lines displayed
        workbook.Save("output.html", htmlOptions);
    }
}
