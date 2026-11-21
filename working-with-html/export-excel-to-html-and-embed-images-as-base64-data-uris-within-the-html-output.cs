// Title: How to export an Excel workbook to HTML with embedded Base64 images using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# export Excel to HTML with images embedded as Base64 | Save workbook as HTML with inline images using Aspose.Cells .NET | Convert .xlsx to HTML and embed pictures as data URIs in C# | How to embed worksheet images in HTML output with Aspose.Cells | Base64 image embedding when saving Excel as HTML with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | HTML export with embedded Base64 images | Excel to HTML conversion with inline pictures | C# Aspose.Cells save workbook as HTML | Base64 data URI image embedding in HTML output

using System;
using Aspose.Cells;

// The example loads an Excel file, configures HtmlSaveOptions to embed worksheet images as Base64 data URIs, and saves the workbook as an HTML file, producing a self‑contained HTML document with inline images.
class Program
{
    static void Main()
    {
        // Load the Excel workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to embed images as Base64 data URIs
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportImagesAsBase64 = true;          // Embed images directly in HTML
        htmlOptions.ExportActiveWorksheetOnly = false;   // Export all worksheets (optional)

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
