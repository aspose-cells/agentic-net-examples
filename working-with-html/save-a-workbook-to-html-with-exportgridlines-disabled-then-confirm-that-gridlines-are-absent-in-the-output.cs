// Title: Save Aspose.Cells Workbook as HTML without Gridlines and Verify Absence (C#)
// Description: Shows how to create a workbook, enable its gridlines, add sample data, set HtmlSaveOptions.ExportGridLines = false, export to HTML, and programmatically confirm that no gridline or border CSS is present in the generated file.
// Keywords: Aspose.Cells | C# HTML export | ExportGridLines false | disable gridlines | verify HTML output | gridline removal | HtmlSaveOptions | Excel to HTML conversion | no border styles | Aspose.Cells sample code
// Common Searches: Aspose.Cells export to HTML without gridlines | C# HtmlSaveOptions ExportGridLines example | how to hide gridlines in HTML output using Aspose.Cells | verify that HTML export has no border CSS | disable gridlines when saving workbook as HTML
// Developer Intent: Export a workbook to HTML with gridlines turned off and programmatically confirm they are omitted.
// Use Cases: Generate clean web reports from Excel data where gridlines would clutter the layout. | Create printable HTML versions of spreadsheets that match corporate branding without visible borders. | Automate a CI check to ensure the ExportGridLines setting is respected in automated HTML exports.
// AI Prompts: Write C# code using Aspose.Cells to save a workbook as HTML with ExportGridLines set to false and validate that the output contains no border or gridline CSS. | Provide a C# unit test that asserts the generated HTML from Aspose.Cells does not include any gridline styles when ExportGridLines is disabled. | Explain how to combine HtmlSaveOptions settings for custom HTML export while keeping gridlines hidden.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to create a workbook, enable its gridlines, add sample data, set HtmlSaveOptions.ExportGridLines = false, export to HTML, and programmatically confirm that no gridline or border CSS is present in the generated file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Enable gridlines in the worksheet (they are visible in Excel)
        worksheet.IsGridlinesVisible = true;

        // Add some sample data so the HTML has content to render
        worksheet.Cells["A1"].PutValue("Sample Text");
        worksheet.Cells["B2"].PutValue(12345);
        worksheet.Cells["C3"].PutValue(DateTime.Now);

        // Configure HTML save options with ExportGridLines disabled
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            ExportGridLines = false   // Ensure gridlines are not exported
        };

        // Define the output HTML file path
        string outputPath = "output.html";

        // Save the workbook as HTML using the configured options
        workbook.Save(outputPath, htmlOptions);

        // Read the generated HTML file
        string htmlContent = File.ReadAllText(outputPath);

        // Simple verification: check that the HTML does not contain border styles
        bool gridLinesFound = htmlContent.Contains("border") || htmlContent.Contains("gridline");

        // Output verification results
        Console.WriteLine($"ExportGridLines option set to: {htmlOptions.ExportGridLines}");
        Console.WriteLine($"Gridlines present in HTML? {(gridLinesFound ? "Yes" : "No")}");
    }
}
