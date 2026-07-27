// Title: C# – Export Excel to HTML with Aspose.Cells, Disable Default CSS and Use Custom TableCssId
// Description: Demonstrates how to save an Aspose.Cells workbook as HTML without the library‑generated CSS. By setting HtmlSaveOptions.DisableCss = true and assigning a custom HtmlSaveOptions.TableCssId, the <table> tag receives a unique ID that can be styled with your own external stylesheet.
// Keywords: Aspose.Cells | HtmlSaveOptions.DisableCss | TableCssId | C# HTML export | Excel to HTML without CSS | .NET Aspose.Cells example | custom table styling
// Common Searches: Aspose.Cells disable default CSS when exporting to HTML | How to set TableCssId for HTML tables in Aspose.Cells C# | Export Excel workbook to HTML without generated CSS | Aspose.Cells HTML export custom table ID | C# Aspose.Cells HTMLSaveOptions example
// Developer Intent: Turn off Aspose.Cells’ built‑in CSS output and rely on a custom TableCssId to apply external styling to the exported HTML table.
// Use Cases: Create lightweight HTML reports that inherit site‑wide CSS frameworks via a unique table ID. | Generate HTML snippets for web applications that manage all styling through their own stylesheet, avoiding Aspose.Cells‑generated styles. | Produce HTML email content where only minimal external CSS is allowed, using TableCssId to target the table.
// AI Prompts: Show C# code that exports an Aspose.Cells workbook to HTML with DisableCss=true and a custom TableCssId, then link an external CSS file for styling. | Provide CSS rules that target the TableCssId set in HtmlSaveOptions to style borders, fonts, and colors of the exported table. | Explain the interaction between DisableCss and TableCssId in Aspose.Cells and how they enable HTML that depends solely on external CSS.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to save an Aspose.Cells workbook as HTML without the library‑generated CSS. By setting HtmlSaveOptions.DisableCss = true and assigning a custom HtmlSaveOptions.TableCssId, the <table> tag receives a unique ID that can be styled with your own external stylesheet.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Alice");
            sheet.Cells["B3"].PutValue(25);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Disable generation of external CSS files; use only inline styles
            htmlOptions.DisableCss = true;

            // Assign a custom TableCssId prefix; this will be added to the <table> element
            // and can be used in external CSS to style the table without relying on default CSS.
            htmlOptions.TableCssId = "custom-table";

            // Save the workbook as HTML using the configured options
            workbook.Save("ExportedTable.html", htmlOptions);
        }
    }
}
