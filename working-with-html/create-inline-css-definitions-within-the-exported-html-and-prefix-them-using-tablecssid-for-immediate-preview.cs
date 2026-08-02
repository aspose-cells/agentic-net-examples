// Title: C# – Export Excel to a Single HTML File with Inline CSS Using TableCssId (Aspose.Cells)
// Description: Demonstrates how to create a workbook, fill it with data, and save it as a single HTML file that embeds custom CSS inside a <style> tag. The TableCssId property prefixes all table selectors, enabling immediate preview without external style sheets.
// Keywords: Aspose.Cells | C# HTML export | inline CSS | TableCssId | single HTML file | embed CSS in HTML | Excel to HTML | custom table styling
// Common Searches: Aspose.Cells embed CSS in HTML export | TableCssId prefix for HTML tables C# | save workbook as single HTML file Aspose | inline CSS with Aspose.Cells HtmlSaveOptions | custom CSS for exported Excel HTML
// Developer Intent: Export an Excel workbook to a self‑contained HTML page that includes custom inline CSS prefixed by a TableCssId for instant visual preview.
// Use Cases: Generate web‑ready reports from Excel with all styling bundled in one file. | Create preview pages for spreadsheets where external CSS files are undesirable. | Maintain consistent table appearance across multiple exported HTML documents by reusing the same TableCssId and CSS definitions.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to HTML, embedding CSS that uses a custom TableCssId prefix. | Show how to adjust the CssStyles property to change header background color and cell padding while keeping SaveAsSingleFile enabled. | Explain the effect of TableCssId on generated HTML selectors and how to reference them in inline CSS.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExportDemo
{
    // Demonstrates how to create a workbook, fill it with data, and save it as a single HTML file that embeds custom CSS inside a <style> tag. The TableCssId property prefixes all table selectors, enabling immediate preview without external style sheets.
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
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Prefix for CSS class names applied to table elements (tr, td, etc.)
            htmlOptions.TableCssId = "custom-table";

            // Enable single‑file output so that CssStyles are embedded in the HTML <style> tag
            htmlOptions.SaveAsSingleFile = true;

            // Define inline CSS that will be placed inside the generated HTML.
            // The CSS selectors use the TableCssId prefix defined above.
            htmlOptions.CssStyles = @"
                .custom-table tr { background-color:#f9f9f9; }
                .custom-table td { border:1px solid #ddd; padding:5px; }
                .custom-table th { background-color:#e0e0e0; font-weight:bold; }
            ";

            // Save the workbook as an HTML file with the specified options
            workbook.Save("ExportedWithInlineCss.html", htmlOptions);

            Console.WriteLine("HTML file generated with inline CSS and TableCssId prefix.");
        }
    }
}
