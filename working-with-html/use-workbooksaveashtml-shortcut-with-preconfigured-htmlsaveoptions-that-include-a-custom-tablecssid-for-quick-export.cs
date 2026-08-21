// Title: Export Workbook to HTML with Custom TableCssId using AspNet.Cells HtmlSaveOptions (C#)
// Description: Shows how to create a Workbook, assign a custom TableCssId via HtmlSaveOptions, and save the file as HTML in one step, making it simple to apply external CSS to the generated table.
// Keywords: Aspose.Cells | HtmlSaveOptions | TableCssId | C# HTML export | custom table id | Excel to HTML | save as HTML Aspose.Cells | HTML export options | CSS styling Aspose.Cells | quick HTML export
// Common Searches: Aspose.Cells set TableCssId when exporting to HTML | C# save workbook as HTML with custom table id | HtmlSaveOptions shortcut for HTML export | apply CSS ID to Aspose.Cells HTML table | export Excel to HTML using Aspose.Cells
// Developer Intent: Export a workbook to HTML while assigning a custom CSS identifier to the table.
// Use Cases: Link an external stylesheet to style the exported table consistently. | Generate multiple HTML reports that share a common table ID for JavaScript manipulation. | Automate one‑line HTML export in batch processes with pre‑configured options.
// AI Prompts: Write C# code that creates HtmlSaveOptions with a custom TableCssId and saves a workbook to HTML, then references an external CSS file. | Explain how TableCssId influences the HTML output and how to target the table with CSS or JavaScript after export. | Show how to export each worksheet to separate HTML files, assigning a unique TableCssId to each for individualized styling.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Shows how to create a Workbook, assign a custom TableCssId via HtmlSaveOptions, and save the file as HTML in one step, making it simple to apply external CSS to the generated table.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.2);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.8);

            // Configure HTML save options with a custom TableCssId
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            htmlOptions.TableCssId = "custom-table-style";

            // Save the workbook as HTML using the pre‑configured options
            workbook.Save("Exported.html", htmlOptions);

            Console.WriteLine("Workbook exported to HTML with TableCssId = " + htmlOptions.TableCssId);
        }
    }
}
