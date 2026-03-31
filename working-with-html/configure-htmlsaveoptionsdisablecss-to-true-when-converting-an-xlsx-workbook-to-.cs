using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing XLSX workbook from disk
            // Replace "input.xlsx" with the actual path to your source file
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Set DisableCss to true so that only inline styles are used
            // and no external or embedded CSS is generated
            htmlOptions.DisableCss = true;

            // Save the workbook as an HTML file using the configured options
            // The output file will contain styling applied directly to each element
            workbook.Save("output.html", htmlOptions);
        }
    }
}