// Title: Export a styled worksheet to HTML with embedded CSS and Base64 images using Aspose.Cells for C#
// AI Prompts: Write C# code that uses Aspose.Cells to save a workbook as an HTML file with inline CSS so that all cell formatting (fonts, colors, patterns) is retained. | Show how to enable HtmlSaveOptions.ExportImagesAsBase64 in Aspose.Cells to embed worksheet images directly as Base64 strings in the generated HTML.
// Common Searches: C# Aspose.Cells export worksheet to HTML with inline CSS preserving cell styles | How to embed images as Base64 in HTML output using Aspose.Cells HtmlSaveOptions | Save Excel workbook as HTML with embedded CSS using Aspose.Cells .NET | Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 example in C# | Retain background and font colors when converting Excel to HTML with Aspose.Cells
// Tags: Aspose.Cells HTML export inline CSS | C# Aspose.Cells export with cell styling | Base64 image embedding Aspose.Cells HTML | Preserve Excel formatting in HTML output | Save workbook as HTML using Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;

// The example creates a workbook, applies styling to a cell, configures HtmlSaveOptions to embed CSS and optionally encode images as Base64, and saves the worksheet as an HTML file that retains the original Excel cell formatting and embeds images directly in the markup.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some data with styling
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleData";

            // Example cell with value and style
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Hello Aspose.Cells");

            // Apply a simple style to retain in HTML
            Style style = cell.GetStyle();
            style.Font.Color = Color.White;
            style.ForegroundColor = Color.Blue;
            style.Pattern = BackgroundType.Solid;
            cell.SetStyle(style);

            // Configure HTML save options to embed CSS (and optionally images) directly in the HTML file
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Embed images as Base64 strings (optional)
            htmlOptions.ExportImagesAsBase64 = true;

            // Save the workbook as an HTML file with embedded CSS
            workbook.Save("WorksheetExport.html", htmlOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
