using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX file
            string inputPath = "input.xlsx";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable export of row and column headings (obsolete property used as requested)
            htmlOptions.ExportHeadings = true;

            // Path for the output HTML file
            string outputPath = "output.html";

            // Save the workbook as HTML with the specified options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine("Workbook has been saved to HTML with headings enabled.");
        }
    }
}