using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook from disk
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options and disable downlevel-revealed conditional comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.DisableDownlevelRevealedComments = true;

            // Save the workbook as an HTML file using the configured options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook has been saved to HTML with DisableDownlevelRevealedComments = true.");
        }
    }
}