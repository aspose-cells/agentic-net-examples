// Title: Set HTML Image DPI to 150 and Export Sharper HTML with Aspose.Cells for .NET
// Description: Shows how to use Aspose.Cells for .NET to set HtmlSaveOptions.ImageOptions.HorizontalResolution and VerticalResolution to 150 DPI, turn off Base64 image embedding, and save a workbook as an HTML file with external high‑resolution images.
// Keywords: Aspose.Cells HTML DPI 150 | C# HtmlSaveOptions ImageOptions | export Excel to HTML high resolution | 150 DPI images Aspose.Cells | .NET workbook to HTML external images | increase image quality in HTML export | chart rendering DPI Aspose | web publishing Excel images | printing‑ready HTML from Excel
// Common Searches: how to change image DPI when saving Excel as HTML using Aspose.Cells | Aspose.Cells C# export HTML with 150 DPI images | set horizontal and vertical resolution for HTML images in Aspose.Cells | disable Base64 images in Aspose.Cells HTML export | increase image sharpness in Excel to HTML conversion
// Developer Intent: Configure the HTML export options so that all rendered images use a 150 DPI resolution, producing clearer graphics in the generated HTML file.
// Use Cases: Creating web‑ready reports where charts must meet a 150 DPI printing standard. | Publishing Excel‑based dashboards online with external image files for faster caching and better quality. | Generating documentation that requires high‑resolution images for print‑friendly PDFs generated from HTML.
// AI Prompts: Provide C# code that sets HtmlSaveOptions.ImageOptions.HorizontalResolution and VerticalResolution to 150 DPI and saves the workbook as HTML with external images. | Explain the impact of ImageOrPrintOptions.Dpi on image sharpness in Aspose.Cells HTML output and how to disable Base64 embedding. | Show a step‑by‑step guide to export an Excel workbook to HTML with 150 DPI images for better web publishing.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to use Aspose.Cells for .NET to set HtmlSaveOptions.ImageOptions.HorizontalResolution and VerticalResolution to 150 DPI, turn off Base64 image embedding, and save a workbook as an HTML file with external high‑resolution images.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample HTML Export with higher DPI");
        sheet.Cells["A2"].PutValue("Images rendered at 150 DPI");

        // Initialize HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Access the ImageOrPrintOptions through the ImageOptions property
        ImageOrPrintOptions imageOpts = saveOptions.ImageOptions;

        // Set both horizontal and vertical DPI to 150 for sharper images
        imageOpts.HorizontalResolution = 150;
        imageOpts.VerticalResolution = 150;

        // Export images as separate files (optional, improves readability)
        saveOptions.ExportImagesAsBase64 = false;

        // Save the workbook as HTML using the configured options
        workbook.Save("output.html", saveOptions);
    }
}
