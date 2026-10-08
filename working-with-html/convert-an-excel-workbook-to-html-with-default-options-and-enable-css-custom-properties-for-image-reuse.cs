// Title: Convert an Excel workbook to HTML with external image files and CSS custom properties using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an XLSX file with Aspose.Cells and saves it as HTML while exporting images as separate files for CSS reuse. | Show how to configure HtmlSaveOptions in Aspose.Cells to disable ExportImagesAsBase64 and keep default HTML settings. | Provide a minimal example that converts a workbook to HTML and stores images in the output folder for use with CSS custom properties.
// Common Searches: Aspose.Cells save workbook as HTML with images saved as separate files | C# export Excel to HTML without embedding images as Base64 | How to use CSS custom properties for image reuse in Aspose.Cells HTML output | HtmlSaveOptions ExportImagesAsBase64 false example in .NET | Generate HTML from XLSX using Aspose.Cells and keep images external
// Tags: Aspose.Cells HtmlSaveOptions ExportImagesAsBase64 false | C# export Excel to HTML external images | HTML output with CSS custom properties Aspose.Cells | convert workbook to HTML default settings .NET | separate image files Aspose.Cells HTML export

using System;
using Aspose.Cells;

// Loads input.xlsx, creates HtmlSaveOptions with default settings, sets ExportImagesAsBase64 to false so images are written as separate files, and saves the workbook as output.html, enabling CSS custom property image reuse.
class Program
{
    static void Main()
    {
        // Load the Excel workbook from a file (default load options are used)
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options with default settings
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

        // Disable embedding images as Base64 so that images are saved as separate files.
        // This allows the generated HTML to reference images via CSS custom properties,
        // enabling image reuse across the document.
        htmlOptions.ExportImagesAsBase64 = false;

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}
