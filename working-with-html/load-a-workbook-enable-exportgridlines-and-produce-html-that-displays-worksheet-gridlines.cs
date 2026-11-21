// Title: Generate HTML from an Excel workbook with gridlines using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export worksheet to HTML with grid lines visible | How to enable gridlines when saving Excel as HTML with Aspose.Cells | C# HtmlSaveOptions ExportGridLines property example | Save Excel file as HTML preserving cell borders using Aspose.Cells | Aspose.Cells HTML export showing Excel gridlines in .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines | HTML export of Excel workbook with gridlines in .NET | C# generate HTML from Excel preserving gridlines | configure Aspose.Cells HTML options for gridline visibility | export worksheet to HTML with visible gridlines using Aspose

using System;
using Aspose.Cells;

// The code loads an existing workbook, sets HtmlSaveOptions.ExportGridLines to true, and saves the file as HTML, producing output that displays the original Excel worksheet's gridlines.
class Program
{
    static void Main()
    {
        // Load the existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to export grid lines
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportGridLines = true; // Enable grid lines in the HTML output

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
