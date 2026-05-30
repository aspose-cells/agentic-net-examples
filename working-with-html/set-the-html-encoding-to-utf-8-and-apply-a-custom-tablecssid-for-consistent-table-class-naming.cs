using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsHtmlExportDemo
{
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

            // Create HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Set the encoding to UTF-8 (default is UTF-8, but we set it explicitly)
            saveOptions.Encoding = Encoding.UTF8;

            // Apply a custom TableCssId to prefix CSS class names for table elements
            saveOptions.TableCssId = "custom-table-style";

            // Save the workbook as an HTML file using the configured options
            workbook.Save("output.html", saveOptions);

            Console.WriteLine("HTML file saved with UTF-8 encoding and TableCssId = " + saveOptions.TableCssId);
        }
    }
}