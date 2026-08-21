// Title: Batch Convert Excel Workbooks to HTML with Unique TableCssId using Aspose.Cells for .NET
// Description: A C# console app that scans a folder for .xls/.xlsx files, loads each workbook with Aspose.Cells, assigns a distinct TableCssId (e.g., table1, table2) via HtmlSaveOptions, and saves the result as HTML in a separate output directory, preventing CSS class collisions across files.
// Keywords: Aspose.Cells HTML export | C# batch Excel to HTML | TableCssId | unique table CSS id | convert multiple workbooks | .NET Excel to HTML | avoid CSS conflicts Aspose
// Common Searches: Aspose.Cells set different TableCssId for each HTML export | batch convert Excel files to HTML C# Aspose | unique table CSS identifier when exporting workbooks | export multiple Excel workbooks to HTML without style overlap | C# code to assign sequential TableCssId in Aspose.Cells
// Developer Intent: Automatically convert every Excel workbook in a directory to an HTML file, giving each output a unique TableCssId to keep table styles isolated.
// Use Cases: Generate independent HTML reports for a collection of spreadsheets, each with its own CSS namespace. | Prepare HTML versions of financial models for web portals where table styles must not interfere with one another. | Automate bulk conversion of Excel dashboards to email‑ready HTML pages with distinct table identifiers.
// AI Prompts: Write C# code that uses Aspose.Cells to batch export all Excel files in a folder to HTML, assigning a sequential TableCssId to each file. | Show how to modify the sample to export each worksheet's CSS to separate files and customize the CSS filenames. | Suggest a method to log each conversion (source path, output HTML, TableCssId) to a CSV while preserving the unique identifiers.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchHtmlExport
{
    // A C# console app that scans a folder for .xls/.xlsx files, loads each workbook with Aspose.Cells, assigns a distinct TableCssId (e.g., table1, table2) via HtmlSaveOptions, and saves the result as HTML in a separate output directory, preventing CSS class collisions across files.
    class Program
    {
        static void Main()
        {
            // Directory containing source Excel workbooks
            string sourceDir = @"C:\InputWorkbooks";

            // Directory where HTML files will be saved
            string outputDir = @"C:\HtmlOutputs";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputDir);

            // Get all Excel files in the source directory (supports .xls and .xlsx)
            string[] excelFiles = Directory.GetFiles(sourceDir, "*.*", SearchOption.TopDirectoryOnly);
            
            int index = 1; // Counter to generate distinct TableCssId values

            foreach (string excelPath in excelFiles)
            {
                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Configure HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions();
                
                // Assign a unique TableCssId for each workbook (e.g., "table1", "table2", ...)
                saveOptions.TableCssId = $"table{index}";

                // Optional: export each worksheet's CSS separately to avoid style conflicts
                // saveOptions.ExportWorksheetCSSSeparately = true;

                // Determine output HTML file name (same as workbook name with .html extension)
                string htmlFileName = Path.GetFileNameWithoutExtension(excelPath) + ".html";
                string htmlPath = Path.Combine(outputDir, htmlFileName);

                // Save the workbook as HTML using the configured options
                workbook.Save(htmlPath, saveOptions);

                Console.WriteLine($"Saved '{excelPath}' as HTML with TableCssId='{saveOptions.TableCssId}' to '{htmlPath}'");

                index++;
            }

            Console.WriteLine("Batch export completed.");
        }
    }
}
