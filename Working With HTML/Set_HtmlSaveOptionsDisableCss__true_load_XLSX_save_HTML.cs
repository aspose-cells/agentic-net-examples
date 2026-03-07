using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the source XLSX workbook
            // Replace "input.xlsx" with the path to your Excel file
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Disable external CSS generation; use only inline styles
            htmlOptions.DisableCss = true;

            // Save the workbook as an HTML file using the configured options
            // Replace "output.html" with the desired output path
            workbook.Save("output.html", htmlOptions);
        }
    }
}