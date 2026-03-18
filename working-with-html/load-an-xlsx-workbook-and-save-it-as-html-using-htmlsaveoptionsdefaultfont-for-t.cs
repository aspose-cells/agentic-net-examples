using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX workbook
            string inputPath = "input.xlsx";

            // Path for the resulting HTML file
            string outputPath = "output.html";

            // Load the workbook from the specified XLSX file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Set the default font name to be used when a cell's font is missing
            saveOptions.DefaultFontName = "Arial";

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved as HTML to '{outputPath}' with default font '{saveOptions.DefaultFontName}'.");
        }
    }
}