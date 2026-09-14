// Title: Create a self‑contained HTML file with embedded Base64 images from an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook and saves it as HTML with HtmlSaveOptions.ExportImagesAsBase64 set to true. | Show how to configure Aspose.Cells HtmlSaveOptions to embed all worksheet images as Base64 strings in the generated HTML. | Adapt an existing Aspose.Cells HTML export to produce a single HTML file that contains no external image files.
// Common Searches: Aspose.Cells C# export Excel to HTML with images embedded as Base64 | How to generate a single HTML file from .xlsx using Aspose.Cells without external image files | HtmlSaveOptions ExportImagesAsBase64 example for .NET developers | Embedding worksheet pictures as Base64 in HTML output with Aspose.Cells | Self‑contained HTML export from Excel workbook using Aspose.Cells C#
// Tags: HtmlSaveOptions ExportImagesAsBase64 | embed worksheet images as Base64 Aspose.Cells | export Excel to single HTML file C# | Aspose.Cells HTML conversion with embedded images | self-contained HTML output from .xlsx

using System;
using Aspose.Cells;

// The example loads an Excel workbook, enables HtmlSaveOptions.ExportImagesAsBase64, and saves the workbook as a single HTML file where all images are embedded as Base64 strings, eliminating external image dependencies.
class HtmlExportWithBase64Images
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to embed images as Base64 strings
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
        {
            ExportImagesAsBase64 = true   // Embed images directly in the HTML
        };

        // Save the workbook as a self‑contained HTML file
        workbook.Save("output.html", htmlOptions);
    }
}
