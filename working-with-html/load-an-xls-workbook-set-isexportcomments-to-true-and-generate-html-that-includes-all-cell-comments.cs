using System;
using Aspose.Cells;

namespace AsposeCellsExportComments
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file (XLS format)
            string inputPath = "input.xls";

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options and enable exporting of comments
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                IsExportComments = true   // Include all cell comments in the generated HTML
            };

            // Path for the output HTML file
            string outputPath = "output.html";

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine("HTML file with comments has been generated successfully.");
        }
    }
}