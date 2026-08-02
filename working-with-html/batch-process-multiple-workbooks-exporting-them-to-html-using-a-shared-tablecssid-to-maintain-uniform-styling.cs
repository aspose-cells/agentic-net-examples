// Title: Batch export Excel workbooks to HTML with a common TableCssId – Aspose.Cells C# example
// Description: A C# console utility that walks through a specified directory, loads each .xlsx workbook with Aspose.Cells, and writes an HTML file per workbook. The HtmlSaveOptions.TableCssId property is set to a single identifier so every generated table shares the same CSS selector, enabling consistent styling across all pages.
// Keywords: Aspose.Cells | C# batch Excel to HTML | TableCssId | HtmlSaveOptions | shared CSS id | convert multiple workbooks | export Excel as HTML | console application | folder processing | .NET Excel conversion
// Common Searches: Aspose.Cells convert folder of xlsx files to html | set identical TableCssId for several html exports | batch Excel to HTML C# sample code | apply uniform CSS to Aspose.Cells HTML output | C# script to export all workbooks in a directory as html
// Developer Intent: Transform every Excel file in a given folder into an HTML document that uses one CSS identifier for all tables.
// Use Cases: Publish a collection of financial reports on a website while keeping table styling centralized. | Create HTML snapshots of archived spreadsheets for quick browser preview, with a single CSS hook for site‑wide theming. | Build a web service that receives uploaded Excel files and returns HTML pages ready for client‑side JavaScript manipulation.
// AI Prompts: Show how to add a custom stylesheet link to the generated HTML while still using a common TableCssId. | Extend the batch process to also export chart images alongside each HTML file. | Adapt the loop to handle both .xlsx and .xls formats in the same conversion run.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchWorkbookToHtml
{
    // A C# console utility that walks through a specified directory, loads each .xlsx workbook with Aspose.Cells, and writes an HTML file per workbook. The HtmlSaveOptions.TableCssId property is set to a single identifier so every generated table shares the same CSS selector, enabling consistent styling across all pages.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing the source Excel workbooks
            string sourceFolder = @"C:\InputWorkbooks";

            // Folder where the generated HTML files will be saved
            string outputFolder = @"C:\OutputHtml";

            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Define a shared TableCssId to apply uniform styling across all HTML files
            const string sharedTableCssId = "shared-table-style";

            // Process each .xlsx file in the source folder
            foreach (string excelPath in Directory.GetFiles(sourceFolder, "*.xlsx"))
            {
                // Verify the file still exists before loading
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"File not found (skipped): {excelPath}");
                    continue;
                }

                try
                {
                    // Load the workbook (create rule)
                    Workbook workbook = new Workbook(excelPath);

                    // Configure HTML save options with the shared TableCssId
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions
                    {
                        TableCssId = sharedTableCssId
                    };

                    // Determine the output HTML file name (same as Excel file name)
                    string htmlFileName = Path.GetFileNameWithoutExtension(excelPath) + ".html";
                    string htmlPath = Path.Combine(outputFolder, htmlFileName);

                    // Save the workbook as HTML (save rule)
                    workbook.Save(htmlPath, saveOptions);

                    Console.WriteLine($"Converted '{excelPath}' to '{htmlPath}' using TableCssId='{sharedTableCssId}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{excelPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
