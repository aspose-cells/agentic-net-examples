using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchHtmlExport
{
    class Program
    {
        static void Main(string[] args)
        {
            // Define the folder containing the source Excel files
            string sourceFolder = @"C:\InputWorkbooks";

            // Define the folder where the HTML files will be saved
            string outputFolder = @"C:\OutputHtml";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files (any supported format) from the source folder
            string[] workbookFiles = Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly);
            
            foreach (string workbookPath in workbookFiles)
            {
                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Create HTML save options and enable both gridlines and comments export
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportGridLines = true,      // Export the gridlines
                    IsExportComments = true,     // Export cell comments
                    ExportActiveWorksheetOnly = false // Export the whole workbook (optional)
                };

                // Build the output HTML file name (same base name as the workbook)
                string htmlFileName = Path.GetFileNameWithoutExtension(workbookPath) + ".html";
                string htmlOutputPath = Path.Combine(outputFolder, htmlFileName);

                // Save the workbook as HTML using the configured options
                workbook.Save(htmlOutputPath, htmlOptions);

                Console.WriteLine($"Saved HTML: {htmlOutputPath}");
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}