// Title: How to export an Aspose.Cells worksheet to an HTML file with embedded CSS styling in C#
// AI Prompts: Write C# code that uses Aspose.Cells to generate an HTML document with the stylesheet embedded, so that the original cell formatting (fonts, colors, borders) is retained. | Show how to configure HtmlSaveOptions in Aspose.Cells for .NET to include CSS inside the HTML output instead of creating a separate .css file.
// Common Searches: C# Aspose.Cells export worksheet to HTML with embedded stylesheet | preserve cell formatting when saving Excel as HTML using Aspose.Cells | Aspose.Cells HtmlSaveOptions embed CSS instead of external file | save Excel worksheet as HTML with styles retained in .NET | export Excel to HTML with internal CSS using Aspose.Cells API
// Tags: Aspose.Cells HtmlSaveOptions include CSS in HTML | export worksheet to HTML retaining formatting Aspose.Cells | C# generate HTML from Excel with embedded stylesheet Aspose.Cells | preserve cell styles in HTML export Aspose.Cells | HTML output with internal CSS Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;

// The program creates a workbook, adds sample data, applies a bold blue header style, configures HtmlSaveOptions to embed CSS directly in the HTML, and saves the worksheet as an HTML file on the desktop.
class ExportWorksheetToHtml
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some data
        worksheet.Cells["A1"].PutValue("Name");
        worksheet.Cells["B1"].PutValue("Age");
        worksheet.Cells["A2"].PutValue("John Doe");
        worksheet.Cells["B2"].PutValue(30);

        // Apply a simple style to demonstrate CSS retention
        Style headerStyle = worksheet.Cells["A1"].GetStyle();
        headerStyle.Font.IsBold = true;
        headerStyle.Font.Color = Color.Blue;
        worksheet.Cells["A1"].SetStyle(headerStyle);

        // Configure HTML save options to embed CSS (default behavior)
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportWorksheetCSSSeparately = false; // embed CSS in the HTML file
        saveOptions.DisableCss = false; // ensure CSS is used instead of inline styles only

        // Define output path (e.g., Desktop)
        string outputPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "WorksheetExport.html");

        // Save the workbook as an HTML file with embedded CSS
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine("Worksheet exported to HTML with embedded CSS at:");
        Console.WriteLine(outputPath);
    }
}
