using System;
using Aspose.Cells;

namespace ExportSimilarBorderStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Configure HTML save options to export similar border styles
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportSimilarBorderStyle = true // Enable similar border style export
            };

            // Save the workbook as HTML using the configured options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook has been saved to HTML with similar border style export.");
        }
    }
}