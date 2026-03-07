using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options and disable downlevel-revealed conditional comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.DisableDownlevelRevealedComments = true;

            // Save the workbook as HTML using the configured options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook saved to HTML with downlevel-revealed comments disabled.");
        }
    }
}