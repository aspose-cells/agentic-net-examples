// Title: Export an Aspose.Cells workbook to HTML without generating an external CSS file using HtmlSaveOptions.DisableCss (C#)
// AI Prompts: Write C# code that saves an Aspose.Cells workbook as HTML while turning off external CSS generation. | Demonstrate how to set HtmlSaveOptions.DisableCss = true before calling Workbook.Save for HTML output. | Show a complete example that creates a workbook, adds data, configures HTML save options to suppress CSS, and writes the HTML file.
// Common Searches: Aspose.Cells C# export to HTML without external CSS file | How to stop Aspose.Cells from creating a CSS file when saving as HTML | HtmlSaveOptions.DisableCss property usage example | Generate HTML from Excel in .NET without linked CSS stylesheet
// Tags: Aspose.Cells HtmlSaveOptions.DisableCss | C# HTML export without CSS | disable external CSS Aspose.Cells | Excel to HTML options .NET | suppress CSS file generation Aspose.Cells

using System;
using Aspose.Cells;

// Creates a workbook, fills sample cells, sets HtmlSaveOptions.DisableCss to true to prevent an external CSS file, and saves the workbook as an HTML document.
class HtmlExportExample
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Score");
        sheet.Cells["A2"].PutValue("Alice");
        sheet.Cells["B2"].PutValue(85);
        sheet.Cells["A3"].PutValue("Bob");
        sheet.Cells["B3"].PutValue(92);

        // Configure HTML save options to disable external CSS generation
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.DisableCss = true; // Prevent creation of external CSS file

        // Export the workbook to HTML using the configured options
        workbook.Save("ExportedReport.html", htmlOptions);
    }
}
