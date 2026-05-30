using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data (including header row)
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.20);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.80);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Include row and column headings in the HTML output
            saveOptions.ExportRowColumnHeadings = true;

            // Prefix CSS class names for table elements with a custom identifier
            saveOptions.TableCssId = "custom-table";

            // Save the workbook as HTML using the configured options
            workbook.Save("ExportedWithHeadings.html", saveOptions);
        }
    }
}