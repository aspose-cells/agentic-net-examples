using System;
using Aspose.Cells;

namespace ExportCommentsExample
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options and enable comment export
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                IsExportComments = true // Export cell comments to the HTML output
            };

            // Save the workbook as an HTML file using the configured options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook saved to HTML with comments exported.");
        }
    }
}