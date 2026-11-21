// Title: Export an Excel workbook to HTML with worksheet gridlines rendered as CSS borders using Aspose.Cells for .NET
// Common Searches: Aspose.Cells export Excel to HTML preserving gridlines as CSS borders C# | HtmlSaveOptions ExportGridLines property example in .NET | Save Excel workbook as HTML with gridlines and base64 images using Aspose.Cells | C# convert .xlsx to .html with visible cell borders Aspose | How to keep Excel gridlines when generating HTML with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines | C# Excel to HTML CSS border rendering | Aspose.Cells ExportImagesAsBase64 option | Preserve worksheet gridlines in HTML output | Convert .xlsx to .html using Aspose.Cells .NET

using Aspose.Cells;
using Aspose.Cells.Rendering;

// The code loads an Excel workbook, sets HtmlSaveOptions.ExportGridLines and ExportImagesAsBase64 to true, and saves the file as HTML, causing worksheet gridlines to be emitted as CSS border definitions and images to be embedded as Base64.
class Program
{
    static void Main()
    {
        // Load the source Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to preserve worksheet gridlines as CSS borders
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // When true, gridlines are rendered as CSS border definitions in the HTML output
            ExportGridLines = true,

            // Optional: embed images directly in the HTML using Base64 encoding
            ExportImagesAsBase64 = true
        };

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}
