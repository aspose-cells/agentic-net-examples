// Title: Convert an Excel workbook to HTML with embedded Base64 images using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells and saves it as an HTML file while embedding all worksheet images as data URI strings. | Show how to set the HtmlSaveOptions flag that causes images to be emitted as inline Base64 strings for reuse in the generated HTML. | Provide a minimal example that uses default HtmlSaveOptions settings except for Base64 image export to produce an HTML file from a workbook.
// Common Searches: Aspose.Cells C# export Excel to HTML with images embedded as Base64 data URIs | How to enable image reuse via CSS custom properties when converting Excel to HTML with Aspose.Cells | Saving a workbook as HTML using default options and Base64 images in .NET | Embedding worksheet pictures in HTML output from Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 | C# Excel to HTML conversion with embedded images | default HTML save options Aspose.Cells | CSS custom properties for image reuse Aspose.Cells | save workbook as HTML Aspose.Cells .NET

using Aspose.Cells;

// The program loads an Excel workbook from 'input.xlsx', configures HtmlSaveOptions with ExportImagesAsBase64 set to true to embed worksheet images as Base64 strings, and saves the workbook as 'output.html' using the default HTML conversion settings.
class Program
{
    static void Main()
    {
        // Load the Excel workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options with default settings
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Enable CSS custom properties for image reuse by embedding images as Base64 strings
        htmlOptions.ExportImagesAsBase64 = true;

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
