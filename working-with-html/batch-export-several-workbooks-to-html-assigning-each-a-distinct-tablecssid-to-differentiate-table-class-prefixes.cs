// Title: Batch export multiple Excel workbooks to HTML with a distinct TableCssId for each file using Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates over a list of .xlsx paths, loads each workbook with Aspose.Cells, and saves it as HTML while assigning a unique TableCssId (e.g., table0, table1) via HtmlSaveOptions. | Generate a C# program that batch converts Excel files to HTML, creates a separate output folder, and includes error handling for missing source files and exceptions. | Provide a C# example that demonstrates how to set the TableCssId property in HtmlSaveOptions for each workbook in a loop to produce HTML files with different table CSS identifiers.
// Common Searches: how to set a different TableCssId for each workbook when exporting to HTML with Aspose.Cells | c# batch convert several xlsx files to html using Aspose.Cells and customize table ids | Aspose.Cells HtmlSaveOptions TableCssId unique per file example | export multiple Excel workbooks to HTML with custom table CSS prefixes in .NET | handle missing Excel files while batch converting to HTML using Aspose.Cells
// Tags: batch export excel to html Aspose.Cells | HtmlSaveOptions TableCssId customization | c# loop convert multiple workbooks | error handling missing source files Aspose.Cells | unique table css identifier per html export

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The program loops through a predefined list of Excel file paths, loads each workbook with Aspose.Cells, and saves it as an HTML file in a target directory. For every export it sets a distinct TableCssId (e.g., table0, table1) via HtmlSaveOptions, ensures the output folder exists, skips missing files, and logs any processing errors.
class BatchHtmlExport
{
    static void Main()
    {
        // List of source workbook file paths to be exported.
        List<string> workbookPaths = new List<string>
        {
            @"C:\Data\Report1.xlsx",
            @"C:\Data\Report2.xlsx",
            @"C:\Data\Report3.xlsx"
        };

        // Destination folder for the generated HTML files.
        string outputFolder = @"C:\Data\HtmlExport\";

        // Ensure the output folder exists.
        Directory.CreateDirectory(outputFolder);

        // Process each workbook.
        for (int i = 0; i < workbookPaths.Count; i++)
        {
            string wbPath = workbookPaths[i];

            // Verify that the source file exists.
            if (!File.Exists(wbPath))
            {
                Console.WriteLine($"File not found: '{wbPath}'. Skipping.");
                continue;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(wbPath);

                // Create HTML save options.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Assign a distinct TableCssId for each workbook (e.g., "table0", "table1", etc.).
                    TableCssId = $"table{i}"
                };

                // Determine the output HTML file name.
                string htmlFileName = Path.GetFileNameWithoutExtension(wbPath) + ".html";
                string htmlPath = Path.Combine(outputFolder, htmlFileName);

                // Save the workbook as HTML with the specified options.
                workbook.Save(htmlPath, saveOptions);

                Console.WriteLine($"Exported '{wbPath}' to HTML with TableCssId='{saveOptions.TableCssId}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{wbPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch export completed.");
    }
}
