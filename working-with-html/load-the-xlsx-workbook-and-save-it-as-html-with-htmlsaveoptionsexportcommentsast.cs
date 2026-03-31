using System;
using Aspose.Cells;

namespace ExportCommentsAsTooltips
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX workbook
            string inputPath = "input.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Enable exporting of cell comments as tooltips
            htmlOptions.IsExportComments = true;

            // Path for the resulting HTML file
            string outputPath = "output.html";

            // Save the workbook as HTML with the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine("Workbook saved as HTML with comments exported as tooltips.");
        }
    }
}