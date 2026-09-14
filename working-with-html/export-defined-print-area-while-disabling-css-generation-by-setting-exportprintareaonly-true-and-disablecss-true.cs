// Title: Export only the defined print area to HTML with inline styles using Aspose.Cells for .NET
// AI Prompts: Write C# code that saves a workbook’s print area as an HTML file with all CSS inlined, using Aspose.Cells. | Show how to set HtmlSaveOptions.ExportPrintAreaOnly and HtmlSaveOptions.DisableCss to produce a minimal HTML export. | Demonstrate exporting a specific cell range to HTML without generating external stylesheet files in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# export specific range to HTML inline CSS | How to save only the print area of a worksheet as HTML with Aspose.Cells | Disable external CSS when converting Excel to HTML using Aspose.Cells .NET | HtmlSaveOptions ExportPrintAreaOnly true example code | Export workbook to HTML without separate stylesheet Aspose.Cells
// Tags: html export print area aspnet | disable external css aspose cells | exportprintareaonly htmlsaveoptions | inline style html export aspose cells | c# workbook to html without css file

using System;
using Aspose.Cells;

// The example creates a workbook, defines a print area (A1:B2), configures HtmlSaveOptions with ExportPrintAreaOnly = true and DisableCss = true, and saves the selected range to an HTML file where all styling is written inline, eliminating external CSS files.
class ExportPrintAreaExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some cells with sample data
        sheet.Cells["A1"].PutValue("First");
        sheet.Cells["B1"].PutValue("Second");
        sheet.Cells["A2"].PutValue(123);
        sheet.Cells["B2"].PutValue(456);

        // Define the print area (only cells A1:B2 will be exported)
        sheet.PageSetup.PrintArea = "A1:B2";

        // Configure HTML export options:
        // - ExportPrintAreaOnly = true ensures only the defined print area is exported.
        // - DisableCss = true prevents generation of CSS files; styles are written inline.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions
        {
            ExportPrintAreaOnly = true,
            DisableCss = true
        };

        // Export the workbook to an HTML file using the configured options
        workbook.Save("ExportedPrintArea.html", saveOptions);
    }
}
