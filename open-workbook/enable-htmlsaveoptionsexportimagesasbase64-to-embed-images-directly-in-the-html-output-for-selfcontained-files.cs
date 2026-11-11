// Title: How to embed Excel images as Base64 in a self‑contained HTML file with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook to HTML with images embedded as Base64 | Set HtmlSaveOptions.ExportImagesAsBase64 to true in .NET | Generate self‑contained HTML from an Excel file using Aspose.Cells | Convert Excel to HTML with embedded images using C# Aspose.Cells | Save Excel workbook as HTML with Base64‑encoded images in C#
// Tags: HtmlSaveOptions ExportImagesAsBase64 | self-contained HTML output Aspose.Cells | embed images as Base64 in HTML C# | Excel to HTML conversion Aspose.Cells | C# save workbook as HTML with embedded images

using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads an Excel workbook, configures HtmlSaveOptions to enable ExportImagesAsBase64, and saves the workbook as a self‑contained HTML file where all images are embedded as Base64 strings.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to embed images as Base64 strings
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.ExportImagesAsBase64 = true; // Enable self‑contained HTML

        // Save the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
