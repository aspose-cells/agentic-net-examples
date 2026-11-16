// Title: Export a specific print area to HTML with inline styles using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export only the defined print area to HTML | How to disable CSS file generation when saving a workbook as HTML with Aspose.Cells | Aspose.Cells HtmlSaveOptions ExportPrintAreaOnly true example | Save worksheet as HTML with inline styles using Aspose.Cells .NET | Aspose.Cells C# export print area A1:B3 to HTML without external CSS
// Tags: ExportPrintAreaOnly HtmlSaveOptions Aspose.Cells | DisableCss Aspose.Cells HTML export | print area HTML export C# Aspose.Cells | inline styles HtmlSaveOptions Aspose.Cells | worksheet print area to HTML Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook, sets a print area (A1:B3) on the first worksheet, configures HtmlSaveOptions to export only that print area and embed styles inline by disabling CSS generation, and saves the result as an HTML file.
class ExportPrintAreaExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some data in the worksheet
        sheet.Cells["A1"].PutValue("Header1");
        sheet.Cells["B1"].PutValue("Header2");
        sheet.Cells["A2"].PutValue(10);
        sheet.Cells["B2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);
        sheet.Cells["B3"].PutValue(40);

        // Define the print area (e.g., A1:B3)
        sheet.PageSetup.PrintArea = "A1:B3";

        // Configure HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            // Export only the defined print area
            ExportPrintAreaOnly = true,
            // Disable generation of CSS files (styles will be inline)
            DisableCss = true
        };

        // Export the worksheet to HTML
        workbook.Save("PrintAreaExport.html", saveOptions);
    }
}
