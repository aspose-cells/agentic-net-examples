// Title: Batch Convert Multiple Excel Workbooks to HTML with Unique TableCssId using Aspose.Cells for .NET
// Description: C# program that scans a directory for .xlsx files, loads each workbook with Aspose.Cells, and saves it as an HTML file. A distinct TableCssId is assigned per workbook via HtmlSaveOptions, preventing CSS conflicts when the pages are displayed together.
// Keywords: Aspose.Cells | C# | batch export to HTML | HtmlSaveOptions | TableCssId | convert Excel to HTML | multiple workbooks | unique CSS identifier | folder processing | automated Excel conversion
// Common Searches: Aspose.Cells export multiple Excel files to HTML | C# set TableCssId for each workbook | batch convert .xlsx to .html with unique CSS ids | HtmlSaveOptions TableCssId loop example | how to avoid CSS collisions when exporting Excel to HTML
// Developer Intent: Automatically generate separate HTML files from a collection of Excel workbooks, giving each table its own CSS identifier.
// Use Cases: Publish a series of financial reports on a web portal where each report needs an isolated table style. | Provide users with web‑ready previews of uploaded spreadsheets, ensuring style rules do not interfere across files. | Schedule nightly conversion of dashboard workbooks to HTML for intranet distribution, with independent CSS selectors for custom theming.
// AI Prompts: Generate C# code that reads all .xlsx files in a folder and saves each as HTML using Aspose.Cells, assigning a unique TableCssId based on the file name. | Explain the effect of the TableCssId property in HtmlSaveOptions on the resulting HTML markup and how to target it with external CSS. | Extend the sample to add a <link> tag for a custom stylesheet in every exported HTML file while preserving distinct TableCssId values.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchHtmlExport
{
    // C# program that scans a directory for .xlsx files, loads each workbook with Aspose.Cells, and saves it as an HTML file. A distinct TableCssId is assigned per workbook via HtmlSaveOptions, preventing CSS conflicts when the pages are displayed together.
    class Program
    {
        static void Main()
        {
            // Define the folder containing the source Excel files
            string sourceFolder = @"C:\ExcelFiles";

            // Define the folder where the HTML files will be saved
            string outputFolder = @"C:\HtmlExports";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Verify source directory exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Get all Excel files in the source folder (you can adjust the search pattern as needed)
            string[] excelFiles = Directory.GetFiles(sourceFolder, "*.xlsx");

            if (excelFiles.Length == 0)
            {
                Console.WriteLine($"No Excel files found in: {sourceFolder}");
                return;
            }

            // Loop through each workbook and export it to HTML with a unique TableCssId
            for (int i = 0; i < excelFiles.Length; i++)
            {
                string excelPath = excelFiles[i];

                // Ensure the file exists before attempting to load
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"File not found: {excelPath}");
                    continue;
                }

                try
                {
                    // Load the workbook from the file
                    Workbook workbook = new Workbook(excelPath);

                    // Create HTML save options
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions
                    {
                        // Assign a distinct TableCssId for this workbook (e.g., "tableStyle0", "tableStyle1", ...)
                        TableCssId = $"tableStyle{i}"
                    };

                    // Build the output HTML file path (same name as the Excel file but with .html extension)
                    string htmlFileName = Path.GetFileNameWithoutExtension(excelPath) + ".html";
                    string htmlPath = Path.Combine(outputFolder, htmlFileName);

                    // Save the workbook as HTML using the configured options
                    workbook.Save(htmlPath, saveOptions);

                    Console.WriteLine($"Exported '{excelPath}' to '{htmlPath}' with TableCssId = '{saveOptions.TableCssId}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{excelPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch export completed.");
        }
    }
}
