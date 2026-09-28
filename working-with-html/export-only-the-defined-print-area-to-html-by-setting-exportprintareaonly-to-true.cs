// Title: Export only the defined print area of an Excel workbook to HTML with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, sets HtmlSaveOptions.ExportPrintAreaOnly to true, and saves the workbook as an HTML document using Aspose.Cells. | Show how to configure Aspose.Cells HtmlSaveOptions to export just the worksheet's print area when converting Excel to HTML in a .NET application.
// Common Searches: asp.net export excel print area to html using aspose.cells | c# Aspose.Cells HtmlSaveOptions ExportPrintAreaOnly example | how to save only the defined print area of a worksheet as html with Aspose.Cells | convert workbook to html limited to print area Aspose.Cells C#
// Tags: Aspose.Cells HtmlSaveOptions ExportPrintAreaOnly | export Excel print area to HTML C# | convert worksheet print area to HTML Aspose.Cells | C# Aspose.Cells HTML conversion print area

using System;
using Aspose.Cells;

// Loads 'input.xlsx', enables HtmlSaveOptions.ExportPrintAreaOnly to restrict output to the defined print area, and saves the result as 'output.html' using Aspose.Cells.
class ExportPrintAreaToHtml
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and enable exporting only the defined print area
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportPrintAreaOnly = true; // Export only the print area

        // Save the workbook as an HTML file using the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
