using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file (XLSX)
            string inputPath = "input.xlsx";

            // Load the workbook from the XLSX file
            // This uses the Workbook constructor that accepts a file path.
            Workbook workbook = new Workbook(inputPath);

            // Create HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Example option: set the HTML version to HTML5
            saveOptions.HtmlVersion = HtmlVersion.Html5;

            // Example option: save the entire workbook as a single HTML file
            saveOptions.SaveAsSingleFile = true;

            // Example option: include grid lines in the output
            saveOptions.ExportGridLines = true;

            // Define the output HTML file path
            string outputPath = "output.html";

            // Save the workbook as HTML using the specified options
            workbook.Save(outputPath, saveOptions);

            // Verify that the file was created
            if (File.Exists(outputPath))
            {
                Console.WriteLine($"HTML file successfully saved to: {Path.GetFullPath(outputPath)}");
            }
            else
            {
                Console.WriteLine("Failed to save the HTML file.");
            }
        }
    }
}