using System;
using Aspose.Cells;

namespace AsposeCellsExportComments
{
    class Program
    {
        static void Main()
        {
            // Load the source Excel workbook (XLSX format)
            Workbook workbook = new Workbook("input.xlsx");

            // Configure HTML save options to export cell comments as tooltips
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // When true, comments are exported and shown as tooltips in the HTML output
                IsExportComments = true
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook successfully converted to HTML with comments as tooltips.");
        }
    }
}