using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Load the source XLSX workbook from disk
            // (Workbook constructor with file path is used as per the lifecycle rule)
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable CSS custom properties to optimize styling and reuse resources
            htmlOptions.EnableCssCustomProperties = true;

            // Save the workbook as an HTML file using the configured options
            // (Workbook.Save method with HtmlSaveOptions follows the provided save rule)
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook has been converted to HTML with CSS custom properties enabled.");
        }
    }
}