// Title: Convert Excel to HTML with CSS Custom Properties using Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, creates default HtmlSaveOptions, enables EnableCssCustomProperties to reuse images via a single base64 CSS variable, and saves the result as an HTML file.
// Keywords: Aspose.Cells | Excel to HTML conversion | HtmlSaveOptions | EnableCssCustomProperties | CSS custom properties | base64 image reuse | C# .NET | workbook export | HTML export with images | image optimization
// Common Searches: Aspose.Cells export Excel to HTML C# | EnableCssCustomProperties example Aspose.Cells | Reuse images with CSS variables in HTML export | Convert .xlsx to HTML with single base64 image | HtmlSaveOptions default settings Aspose.Cells
// Developer Intent: Export an Excel workbook to HTML while activating CSS custom properties so that embedded images are referenced through a single base64 definition.
// Use Cases: Generate web‑ready reports from spreadsheets with minimal image duplication. | Create compact HTML email bodies where images are shared via a CSS variable. | Automate batch conversion of multiple workbooks to consistent, lightweight HTML pages.
// AI Prompts: Show how to specify a custom CSS variable name for the base64 image when EnableCssCustomProperties is true. | Provide code to write the HTML output to a MemoryStream instead of a file while keeping CSS custom properties enabled. | Explain how to combine EnableCssCustomProperties with ExportImagesAsBase64 and other HtmlSaveOptions for advanced HTML export control.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Loads an Excel workbook, creates default HtmlSaveOptions, enables EnableCssCustomProperties to reuse images via a single base64 CSS variable, and saves the result as an HTML file.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Load the workbook from the file system
            Workbook workbook = new Workbook(sourcePath);

            // Create HTML save options with default settings
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable CSS custom properties to allow image reuse via a single base64 definition
            htmlOptions.EnableCssCustomProperties = true;

            // Path for the generated HTML file
            string outputPath = "output.html";

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully converted to HTML with CSS custom properties enabled: {outputPath}");
        }
    }
}
