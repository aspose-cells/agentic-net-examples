// Title: Export an Excel workbook to HTML with embedded Base64 images and custom page margins using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file, sets left, right, top, and bottom margins on the first worksheet, and saves the workbook as an HTML file with all images embedded as Base64 using Aspose.Cells. | Demonstrate how to configure HtmlSaveOptions in Aspose.Cells to embed worksheet images directly in the exported HTML document.
// Common Searches: Aspose.Cells export worksheet to HTML with embedded Base64 images | C# set page margins before saving Excel as HTML using Aspose.Cells | How to embed images in HTML output when converting Excel with Aspose.Cells .NET | HtmlSaveOptions ExportImagesAsBase64 example for Aspose.Cells
// Tags: html export base64 image embedding Aspose.Cells | worksheet page margin configuration Aspose.Cells C# | HtmlSaveOptions ExportImagesAsBase64 usage | embed images in html output Aspose.Cells | set page margins before html conversion Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Loads an Excel workbook, configures left/right/top/bottom margins on the first worksheet, sets HtmlSaveOptions to embed images as Base64, and saves the result as an HTML file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook from disk
            // Replace "input.xlsx" with the path to your source Excel file
            Workbook workbook = new Workbook("input.xlsx");

            // Enable page margins for the first worksheet
            // Margins are specified in inches (default unit)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.LeftMargin = 0.5;   // Left margin
            sheet.PageSetup.RightMargin = 0.5;  // Right margin
            sheet.PageSetup.TopMargin = 0.75;   // Top margin
            sheet.PageSetup.BottomMargin = 0.75; // Bottom margin

            // Configure HTML save options to embed images as Base64 strings
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.ExportImagesAsBase64 = true; // Embed images directly in the HTML

            // Save the workbook as an HTML file with the specified options
            // The resulting HTML will contain embedded images
            workbook.Save("output.html", htmlOptions);
        }
    }
}
