// Title: Export an Excel worksheet to HTML with visible gridlines using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions.ExportGridLines to true, and saves the workbook as an HTML document. | Provide a .NET example that converts a specific worksheet to HTML while preserving the original Excel gridlines using Aspose.Cells. | Write a C# snippet demonstrating how to enable gridlines in the HTML output of a workbook saved with Aspose.Cells.
// Common Searches: asp.net aspose.cells export html with gridlines c# | how to keep Excel gridlines when saving as html using aspose.cells | c# HtmlSaveOptions ExportGridLines true example | convert first worksheet to html showing cell borders aspose.cells | aspose.cells generate html output that displays worksheet gridlines
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines | HTML export of Excel worksheet with gridlines | C# convert workbook to HTML preserving gridlines | Aspose.Cells HTML output cell borders | save workbook as HTML showing gridlines

using System;
using Aspose.Cells;

// The program loads 'input.xlsx', enables gridline export via HtmlSaveOptions.ExportGridLines, and saves the workbook as 'output.html', producing HTML that displays the worksheet's gridlines.
class Program
{
    static void Main()
    {
        // Load the existing workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (optional, can be any worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure HTML save options to export grid lines
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportGridLines = true; // Enable grid lines in the generated HTML

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
