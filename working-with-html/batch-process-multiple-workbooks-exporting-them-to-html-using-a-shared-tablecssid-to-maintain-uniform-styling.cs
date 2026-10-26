// Title: Export multiple Excel workbooks to HTML with a shared TableCssId using Aspose.Cells for .NET
// Common Searches: how to export a list of .xlsx files to HTML with the same table CSS id using Aspose.Cells in C# | batch convert Excel workbooks to HTML while applying a common TableCssId for uniform styling | Aspose.Cells C# export multiple workbooks to HTML with shared TableCssId option | set HtmlSaveOptions.TableCssId for all worksheets during batch HTML export in .NET | C# loop through Excel files and save each as HTML using Aspose.Cells
// Tags: batch export Excel workbooks to HTML using Aspose.Cells | HtmlSaveOptions TableCssId configuration in C# | apply shared CSS ID to generated HTML tables | iterate workbook list for HTML conversion Aspose.Cells | consistent table styling across multiple HTML exports

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchHtmlExport
{
    // The sample iterates over a collection of Excel file paths, creates an output folder, and calls a helper method that loads each workbook, configures HtmlSaveOptions with a common TableCssId and ExportActiveWorksheetOnly = false, then saves the workbook as an HTML file named after the source workbook, ensuring uniform table styling across all exported HTML files.
    class Program
    {
        static void Main(string[] args)
        {
            // Example input: list of workbook file paths
            List<string> workbookPaths = new List<string>
            {
                @"C:\Workbooks\Book1.xlsx",
                @"C:\Workbooks\Book2.xlsx",
                @"C:\Workbooks\Book3.xlsx"
            };

            // Output directory for HTML files
            string outputDirectory = @"C:\Workbooks\HtmlExport";

            // Shared CSS ID for the HTML tables to ensure uniform styling
            string sharedTableCssId = "sharedTableStyle";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputDirectory);

            // Process each workbook
            foreach (string workbookPath in workbookPaths)
            {
                ExportWorkbookToHtml(workbookPath, outputDirectory, sharedTableCssId);
            }

            Console.WriteLine("Batch export completed.");
        }

        /// <param name="workbookPath">Full path to the source workbook.</param>
        /// <param name="outputDir">Directory where the HTML file will be saved.</param>
        /// <param name="tableCssId">The CSS ID to assign to the generated HTML tables.</param>
        private static void ExportWorkbookToHtml(string workbookPath, string outputDir, string tableCssId)
        {
            // Verify the source workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"File not found: {workbookPath}");
                return;
            }

            try
            {
                // Load the workbook from file
                Workbook workbook = new Workbook(workbookPath);

                // Configure HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Use the same CSS ID for all tables to apply uniform styling
                    TableCssId = tableCssId,

                    // Export all worksheets (default behavior; explicitly set for clarity)
                    ExportActiveWorksheetOnly = false
                };

                // Determine the output HTML file name (same as workbook name with .html extension)
                string htmlFileName = Path.GetFileNameWithoutExtension(workbookPath) + ".html";
                string htmlFilePath = Path.Combine(outputDir, htmlFileName);

                // Save the workbook as HTML using the configured options
                workbook.Save(htmlFilePath, saveOptions);
                Console.WriteLine($"Exported: {htmlFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{workbookPath}': {ex.Message}");
            }
        }
    }
}
