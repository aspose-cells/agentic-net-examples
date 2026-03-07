using System;
using Aspose.Cells;

namespace AsposeCellsExportCommentsExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook from disk
            // (Replace "input.xlsx" with the actual path to your workbook)
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options and enable comment export
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export cell comments to the generated HTML file
                IsExportComments = true
            };

            // Save the workbook as an HTML file with comments included
            // (Replace "output.html" with the desired output path)
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook saved to HTML with comments exported successfully.");
        }
    }
}