using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Load the source Excel file (XLSX)
            // Replace "input.xlsx" with the path to your workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Exclude unused styles to reduce the size of the generated HTML
            // This is true by default, but setting it explicitly clarifies intent
            htmlOptions.ExcludeUnusedStyles = true;

            // Save the workbook as an HTML file using the configured options
            // Replace "output.html" with the desired output path
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Excel file has been converted to HTML with unused styles excluded.");
        }
    }
}