using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Load the source XLSX workbook
            // (Assumes the file "input.xlsx" exists in the application directory)
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable exporting of comments so they appear as tooltips in the HTML output
            // The combination of IsExportComments and ExportCommentsType ensures comments are rendered as tooltips.
            htmlOptions.IsExportComments = true;
            htmlOptions.ExportCommentsType = PrintCommentsType.PrintInPlace; // renders comments as tooltips

            // Save the workbook as an HTML file with the configured options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook successfully converted to HTML with comments as tooltips.");
        }
    }
}