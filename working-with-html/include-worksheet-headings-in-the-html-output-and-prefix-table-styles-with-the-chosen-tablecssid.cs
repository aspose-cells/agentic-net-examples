// Title: C# – Export Excel to HTML with Row/Column Headings and Custom TableCssId using Aspose.Cells
// Description: Demonstrates how to create a workbook, fill cells A1:B3, enable ExportRowColumnHeadings, set TableCssId to a custom prefix, and save the sheet as an HTML file with worksheet headings and prefixed CSS classes.
// Keywords: Aspose.Cells HTML export | ExportRowColumnHeadings | TableCssId | custom CSS prefix | C# Excel to HTML | worksheet headings HTML | .NET Aspose.Cells example
// Common Searches: Aspose.Cells include row and column headings in HTML export | Set custom CSS class prefix for tables in Aspose.Cells HTML output | HtmlSaveOptions ExportRowColumnHeadings C# example | Aspose.Cells TableCssId usage | Export Excel worksheet to HTML with headings .NET
// Developer Intent: Generate an HTML representation of a workbook that shows Excel row/column labels and uses a developer‑defined CSS identifier for table styling.
// Use Cases: Produce HTML reports that retain Excel’s A‑Z column and 1‑N row labels for clear data reference. | Embed exported tables into web pages and apply site‑specific styles by targeting the custom TableCssId selector. | Automate bulk conversion of multiple worksheets to HTML while maintaining consistent CSS class naming across all tables.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to HTML, showing row/column headings and using TableCssId="custom-table-style". | Explain how HtmlSaveOptions.ExportRowColumnHeadings and TableCssId affect the generated HTML markup. | Provide a CSS snippet that styles tables whose class names start with the TableCssId "custom-table-style".

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Demonstrates how to create a workbook, fill cells A1:B3, enable ExportRowColumnHeadings, set TableCssId to a custom prefix, and save the sheet as an HTML file with worksheet headings and prefixed CSS classes.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(15);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Include row and column headings (A, B, 1, 2, …) in the HTML output
            htmlOptions.ExportRowColumnHeadings = true;
            // Prefix CSS class names for table elements with a custom identifier
            htmlOptions.TableCssId = "custom-table-style";

            // Save the workbook as HTML using the configured options
            workbook.Save("ExportedWithHeadings.html", htmlOptions);
        }
    }
}
