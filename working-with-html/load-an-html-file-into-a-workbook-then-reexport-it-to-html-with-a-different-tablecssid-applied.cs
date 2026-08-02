// Title: C# – Load HTML into Aspose.Cells Workbook and Save with a Custom TableCssId
// Description: Shows how to import an HTML file into an Aspose.Cells Workbook using HtmlLoadOptions, assign a custom TableCssId via HtmlSaveOptions, and export the workbook back to HTML so the generated <table> element carries the specified CSS identifier.
// Keywords: Aspose.Cells | C# | HtmlLoadOptions | HtmlSaveOptions | TableCssId | custom table id | export workbook to HTML | load HTML into workbook | change HTML table CSS identifier | .NET Excel to HTML | HTML to Excel conversion
// Common Searches: Aspose.Cells set TableCssId when saving HTML | load HTML file into Workbook C# | change table id in exported HTML Aspose.Cells | HtmlSaveOptions TableCssId example | customize HTML table CSS id Aspose.Cells
// Developer Intent: Load an HTML file into a workbook and re‑export it as HTML with a user‑defined TableCssId.
// Use Cases: Align the generated HTML table ID with existing site‑wide CSS rules after converting legacy HTML reports to Excel workbooks. | Prevent ID collisions when exporting multiple worksheets to HTML by assigning unique TableCssId values to each export. | Automate the transformation of web templates into Excel files and back to HTML while preserving project‑specific CSS selectors.
// AI Prompts: Provide a C# example that loads an HTML file into an Aspose.Cells Workbook, sets TableCssId to 'custom-table', and saves the workbook as HTML. | Explain how HtmlLoadOptions and HtmlSaveOptions work together to change the table CSS identifier during HTML export in Aspose.Cells. | Generate robust C# code that loads an HTML workbook, applies a custom TableCssId, handles possible exceptions, and writes the output HTML file.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlTableCssIdDemo
{
    // Shows how to import an HTML file into an Aspose.Cells Workbook using HtmlLoadOptions, assign a custom TableCssId via HtmlSaveOptions, and export the workbook back to HTML so the generated <table> element carries the specified CSS identifier.
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file
            string inputHtmlPath = "input.html";

            // Load the HTML file into a Workbook using HtmlLoadOptions
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(inputHtmlPath, loadOptions);

            // Configure HTML save options with a new TableCssId
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            saveOptions.TableCssId = "custom-table"; // New prefix for table CSS identifiers

            // Path for the exported HTML file
            string outputHtmlPath = "output.html";

            // Save the workbook back to HTML using the configured options
            workbook.Save(outputHtmlPath, saveOptions);

            Console.WriteLine($"HTML file saved to '{outputHtmlPath}' with TableCssId = '{saveOptions.TableCssId}'.");
        }
    }
}
