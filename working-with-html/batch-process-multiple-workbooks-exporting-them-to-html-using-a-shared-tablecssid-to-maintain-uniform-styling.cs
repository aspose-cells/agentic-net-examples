// Title: Export multiple Excel workbooks to HTML with a shared TableCssId using Aspose.Cells for .NET
// Common Searches: C# batch convert Excel files to HTML with a common TableCssId using Aspose.Cells | How to apply the same table CSS identifier to all HTML exports from Aspose.Cells | Aspose.Cells HtmlSaveOptions TableCssId example for processing many workbooks | Automate exporting a folder of .xlsx files to HTML with uniform table styling in .NET | Save Excel workbooks as HTML with embedded images and shared CSS id using Aspose.Cells
// Tags: Aspose.Cells batch HTML export | HtmlSaveOptions TableCssId setting | C# export Excel to HTML with shared CSS id | uniform table styling Aspose.Cells | embed images as Base64 Aspose.Cells HTML

using System;
using System.IO;
using Aspose.Cells;

namespace BatchHtmlExport
{
    // The program scans a specified input folder for Excel workbooks, loads each file with Aspose.Cells, and saves it as an HTML document using HtmlSaveOptions that assign a common TableCssId, embed images as Base64, and consolidate CSS, enabling consistent table styling across all exported HTML files.
    class Program
    {
        static void Main(string[] args)
        {
            // Define the folder that contains the source workbooks.
            // Adjust these paths as needed for your environment.
            string inputFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "InputWorkbooks");
            string outputFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "OutputHtml");

            // Ensure the output directory exists.
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            // Verify that the input directory exists.
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Define a shared TableCssId that will be applied to every exported HTML file.
            // This ID can be referenced in a custom CSS file to give all tables a uniform style.
            const string sharedTableCssId = "shared-table-style";

            // Get all Excel files in the input folder (supporting common extensions).
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            if (excelFiles.Length == 0)
            {
                Console.WriteLine($"No workbook files found in: {inputFolder}");
                return;
            }

            foreach (string workbookPath in excelFiles)
            {
                // Filter by known Excel extensions.
                string extension = Path.GetExtension(workbookPath).ToLowerInvariant();
                if (extension != ".xlsx" && extension != ".xls" && extension != ".xlsb" && extension != ".xlsm")
                {
                    Console.WriteLine($"Skipping unsupported file: {workbookPath}");
                    continue;
                }

                try
                {
                    Console.WriteLine($"Processing workbook: {Path.GetFileName(workbookPath)}");

                    // Load the workbook.
                    Workbook workbook = new Workbook(workbookPath);

                    // Configure HTML save options.
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        // Apply the shared TableCssId to all tables.
                        TableCssId = sharedTableCssId,

                        // Optional: keep CSS in a single file for easier handling.
                        ExportWorksheetCSSSeparately = false,

                        // Optional: embed images as Base64 to keep a single HTML file.
                        ExportImagesAsBase64 = true
                    };

                    // Determine the output HTML file path.
                    string outputFileName = Path.GetFileNameWithoutExtension(workbookPath) + ".html";
                    string outputPath = Path.Combine(outputFolder, outputFileName);

                    // Save the workbook as HTML using the configured options.
                    workbook.Save(outputPath, htmlOptions);

                    Console.WriteLine($"Successfully exported to: {outputPath}");
                }
                catch (Exception ex)
                {
                    // Report any failure but continue processing remaining files.
                    Console.WriteLine($"Error processing '{workbookPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch export completed.");
        }
    }
}
