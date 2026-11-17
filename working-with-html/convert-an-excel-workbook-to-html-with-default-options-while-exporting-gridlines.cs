// Title: Export an Excel workbook to HTML with gridlines using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to export Excel gridlines to HTML | save workbook as HTML with gridlines using Aspose.Cells .NET | Aspose.Cells HtmlSaveOptions default settings export gridlines | convert input.xlsx to output.html preserving gridlines in C# | Aspose.Cells export Excel to HTML with default options and gridlines
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines | C# Excel to HTML conversion with gridlines | default HTML save options Aspose.Cells | export gridlines to HTML using Aspose.Cells | save workbook as HTML C# Aspose

using System;
using Aspose.Cells;

// The code loads an Excel workbook, enables ExportGridLines on HtmlSaveOptions, and saves the file as HTML, producing an output that retains the original gridlines.
class Program
{
    static void Main()
    {
        // Load the existing Excel workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options with default settings
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Enable exporting of gridlines to the HTML output
        htmlOptions.ExportGridLines = true;

        // Save the workbook as an HTML file using the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
