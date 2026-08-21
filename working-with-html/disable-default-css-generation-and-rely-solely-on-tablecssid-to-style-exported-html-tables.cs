// Title: Aspose.Cells .NET: Export Workbook to HTML with Inline Styles and Custom TableCssId
// Description: Creates an in‑memory workbook, fills sample data, and saves it as HTML using HtmlSaveOptions with DisableCss = true and TableCssId = "custom-table", so styling is inline and the table can be targeted by a custom ID.
// Keywords: Aspose.Cells | HtmlSaveOptions | DisableCss | TableCssId | export Excel to HTML | inline styles | .NET | C# | HTML table styling | custom CSS id
// Common Searches: disable default CSS Aspose.Cells HTML export | set TableCssId in Aspose.Cells | export Excel as HTML with only inline styles | Aspose.Cells HtmlSaveOptions example C# | custom table identifier HTML output Aspose
// Developer Intent: Turn off automatic CSS generation and use TableCssId to apply external styles to the exported HTML table.
// Use Cases: Generate lightweight HTML reports where a site‑wide stylesheet targets the table via a known ID. | Create email‑friendly HTML that cannot reference external CSS files, relying on inline styles and a predictable table ID. | Embed the exported table into an existing web page and apply custom CSS rules using the specified TableCssId.
// AI Prompts: Write C# code that exports an Aspose.Cells workbook to HTML with DisableCss enabled and a custom TableCssId. | Show how to craft CSS that selects the TableCssId set by Aspose.Cells during HTML export. | Demonstrate changing the TableCssId value in the example and verifying the generated HTML contains the new identifier.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Creates an in‑memory workbook, fills sample data, and saves it as HTML using HtmlSaveOptions with DisableCss = true and TableCssId = "custom-table", so styling is inline and the table can be targeted by a custom ID.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Alice");
            sheet.Cells["B3"].PutValue(25);

            // Configure HTML save options:
            // - DisableCss = true forces all styling to be inline (no external CSS files)
            // - TableCssId provides a custom identifier that can be used in CSS selectors
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            htmlOptions.DisableCss = true;          // Use only inline styles
            htmlOptions.TableCssId = "custom-table"; // Prefix for CSS classes/ids in the generated table

            // Save the workbook as HTML using the configured options
            workbook.Save("ExportedTable.html", htmlOptions);

            Console.WriteLine("HTML export completed with inline styles and TableCssId set.");
        }
    }
}
