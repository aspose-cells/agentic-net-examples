using System;
using Aspose.Cells;

namespace AsposeCellsHtmlCrossDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook from disk
            // Replace "input.xlsx" with the path to your source workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Configure the cross‑cell string handling.
            // HtmlCrossType.Cross makes the string cross cells and provides
            // the best performance for large HTML files.
            htmlOptions.HtmlCrossStringType = HtmlCrossType.Cross;

            // Save the workbook as an HTML file using the configured options
            // Replace "output.html" with the desired output path
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook has been saved to HTML with HtmlCrossType.Cross.");
        }
    }
}