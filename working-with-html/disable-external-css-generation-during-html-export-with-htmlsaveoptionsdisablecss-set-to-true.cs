// Title: Disable external CSS generation when exporting an Excel workbook to HTML with Aspose.Cells in C#
// Common Searches: C# Aspose.Cells export workbook to HTML without creating external CSS files | How to embed styles directly in HTML using HtmlSaveOptions.DisableCss in Aspose.Cells | Aspose.Cells HtmlSaveOptions.DisableCss property example for inline CSS | Prevent external stylesheet generation during HTML export with Aspose.Cells C# | Generate HTML from Excel with embedded CSS using Aspose.Cells in .NET
// Tags: Aspose.Cells HtmlSaveOptions DisableCss | HTML export embed CSS Aspose.Cells | C# export workbook to HTML without external stylesheet | inline CSS generation Aspose.Cells | disable external CSS Aspose.Cells HTML export

using System;
using Aspose.Cells;

// The example creates a workbook, adds sample data, configures HtmlSaveOptions with DisableCss set to true so that styles are embedded directly in the HTML, and saves the result as ExportedDocument.html.
class HtmlExportExample
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // For demonstration, an empty workbook is created

        // Add some data to the workbook (optional)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
        sheet.Cells["A2"].PutValue(12345);

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            // Disable generation of external CSS files; styles will be embedded directly in the HTML
            DisableCss = true
        };

        // Export the workbook to HTML using the configured options
        workbook.Save("ExportedDocument.html", htmlOptions);
    }
}
