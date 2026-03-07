using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class DisableDownlevelRevealedCommentsDemo
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            // (Replace "input.xlsx" with the path to your source file)
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Turn off downlevel-revealed conditional comments in the generated HTML
            htmlOptions.DisableDownlevelRevealedComments = true;

            // Save the workbook as HTML using the configured options
            // (Replace "output.html" with the desired output path)
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook saved to HTML with downlevel-revealed comments disabled.");
        }
    }
}