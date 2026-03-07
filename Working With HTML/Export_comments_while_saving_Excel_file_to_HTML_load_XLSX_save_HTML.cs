using System;
using Aspose.Cells;

namespace ExportCommentsToHtml
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Configure HTML save options to export comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Enable exporting of cell comments to the generated HTML
                IsExportComments = true
            };

            // Save the workbook as an HTML file with comments included
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook saved to HTML with comments exported successfully.");
        }
    }
}