// Title: Load HTML into Aspose.Cells Workbook and Export with a Custom TableCssId (C#)
// Description: Demonstrates how to import an HTML file into an Aspose.Cells Workbook using default HtmlLoadOptions, assign a custom TableCssId via HtmlSaveOptions, and save the workbook back to HTML so that table elements carry the specified CSS identifier.
// Keywords: Aspose.Cells | C# | HTML to Workbook conversion | HtmlLoadOptions | HtmlSaveOptions | TableCssId | custom table CSS id | export HTML with Aspose.Cells | load HTML file into workbook
// Common Searches: Aspose.Cells set TableCssId when saving HTML | load HTML file into Aspose.Cells workbook C# | change table CSS identifier in exported HTML Aspose.Cells | HtmlSaveOptions custom TableCssId example | convert HTML to workbook and back with Aspose.Cells
// Developer Intent: Load an existing HTML document into an Aspose.Cells Workbook, apply a custom TableCssId, and re‑export the workbook to HTML.
// Use Cases: Standardize table CSS identifiers across generated HTML reports for consistent site styling. | Batch‑process multiple HTML templates, assigning the same TableCssId to each output file. | Override the default table class when exporting a workbook to match a corporate stylesheet.
// AI Prompts: Generate a C# example that loads an HTML file into an Aspose.Cells Workbook, sets TableCssId to "my-table", and saves it as HTML. | Explain the interaction between HtmlLoadOptions and HtmlSaveOptions in Aspose.Cells, focusing on how TableCssId influences the exported HTML. | Suggest additional HtmlSaveOptions (e.g., ExportImagesAsBase64, ExportActiveWorksheetOnly) that can be combined with a custom TableCssId for HTML export.

using System;
using Aspose.Cells;

// Demonstrates how to import an HTML file into an Aspose.Cells Workbook using default HtmlLoadOptions, assign a custom TableCssId via HtmlSaveOptions, and save the workbook back to HTML so that table elements carry the specified CSS identifier.
class HtmlTableCssIdExample
{
    static void Main()
    {
        // Path to the source HTML file
        string inputHtmlPath = "input.html";

        // Path for the resulting HTML file
        string outputHtmlPath = "output.html";

        // Load the HTML file into a workbook.
        // HtmlLoadOptions can be customized if needed; here we use the defaults.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        Workbook workbook = new Workbook(inputHtmlPath, loadOptions);

        // Configure HTML save options with a custom TableCssId.
        // This prefix will be added to the CSS classes of table elements in the output HTML.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.TableCssId = "custom-table-style";

        // Save the workbook back to HTML using the specified options.
        workbook.Save(outputHtmlPath, saveOptions);

        Console.WriteLine($"HTML file saved to '{outputHtmlPath}' with TableCssId = '{saveOptions.TableCssId}'.");
    }
}
