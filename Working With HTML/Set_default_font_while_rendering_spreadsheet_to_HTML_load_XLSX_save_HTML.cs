using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Load the source XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options and set the default font name
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.DefaultFontName = "Arial"; // Font used when a cell's font is missing

            // Save the workbook as HTML using the specified options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook exported to HTML with default font: " + htmlOptions.DefaultFontName);
        }
    }
}