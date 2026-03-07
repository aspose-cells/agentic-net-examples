using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Load the existing Excel workbook (XLSX format)
            // Replace "input.xlsx" with the path to your source file.
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options and set a custom prefix for table CSS IDs.
            // The TableCssId property defines the prefix used for CSS class names
            // such as tr, col, td within the generated <table> element.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            saveOptions.TableCssId = "custom-table-style";

            // Save the workbook as an HTML file using the configured options.
            // Replace "output.html" with the desired output path.
            workbook.Save("output.html", saveOptions);

            Console.WriteLine("Workbook has been exported to HTML with TableCssId prefix: " + saveOptions.TableCssId);
        }
    }
}