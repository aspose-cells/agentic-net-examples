// Title: Aspose.Cells C# – Export Excel to HTML with 200 DPI Images
// Description: Demonstrates how to configure HtmlSaveOptions.ImageOptions to 200 DPI for both axes, optionally embed images as Base64, and save a workbook as a high‑resolution HTML file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells HTML export DPI | C# set image resolution Aspose.Cells | HtmlSaveOptions ImageOptions 200 DPI | high resolution HTML images Excel | embed images Base64 Aspose.Cells | export workbook to HTML C# | Aspose.Cells image quality settings
// Common Searches: Aspose.Cells set image DPI when saving as HTML | C# export Excel to HTML with high‑resolution images | HtmlSaveOptions ImageOptions horizontal vertical DPI | how to embed images as Base64 in Aspose.Cells HTML output | increase image clarity in HTML export Aspose.Cells
// Developer Intent: Configure image DPI to 200 and generate an HTML file with high‑resolution graphics from an Excel workbook.
// Use Cases: Publish web‑ready reports where charts and graphics retain sharpness. | Create single‑file HTML emails with embedded high‑resolution images. | Produce printable HTML versions of spreadsheets that require detailed visuals.
// AI Prompts: Provide C# code to export an Aspose.Cells workbook to HTML with 300 DPI images and Base64 embedding. | Show how to set both horizontal and vertical DPI in HtmlSaveOptions.ImageOptions and save to a custom folder. | Explain strategies to balance high‑resolution HTML images and file size when using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to configure HtmlSaveOptions.ImageOptions to 200 DPI for both axes, optionally embed images as Base64, and save a workbook as a high‑resolution HTML file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("High‑resolution HTML export example");

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Set image DPI to 200 for both horizontal and vertical resolution
            // HtmlSaveOptions.ImageOptions provides the ImageOrPrintOptions object
            htmlOptions.ImageOptions.HorizontalResolution = 200;
            htmlOptions.ImageOptions.VerticalResolution = 200;

            // Optional: embed images as Base64 to keep a single HTML file
            // htmlOptions.ExportImagesAsBase64 = true;

            // Save the workbook as HTML with the specified high‑resolution image settings
            string outputPath = "HighResolutionOutput.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved to '{outputPath}' with image DPI set to 200.");
        }
    }
}
