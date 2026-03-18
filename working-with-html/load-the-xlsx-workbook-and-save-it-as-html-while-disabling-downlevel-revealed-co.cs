using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the existing XLSX workbook
            // Replace "input.xlsx" with the path to your source workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Disable downlevel-revealed conditional comments in the generated HTML
            htmlOptions.DisableDownlevelRevealedComments = true;

            // Save the workbook as HTML using the configured options
            // Replace "output.html" with the desired output file path
            workbook.Save("output.html", htmlOptions);
        }
    }
}