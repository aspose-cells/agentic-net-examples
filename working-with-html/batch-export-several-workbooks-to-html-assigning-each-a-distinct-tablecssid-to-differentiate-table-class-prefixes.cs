// Title: Export multiple Excel workbooks to separate HTML files with unique TableCssId using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells batch convert Excel files to HTML with different TableCssId values | How to assign a distinct TableCssId for each workbook when saving as HTML in Aspose.Cells | Automate exporting a list of .xlsx files to HTML with unique table CSS identifiers in .NET | Aspose.Cells HtmlSaveOptions TableCssId per workbook example | Generate separate HTML tables with custom IDs from multiple Excel workbooks using C#
// Tags: Aspose.Cells HtmlSaveOptions TableCssId | batch Excel to HTML conversion C# | unique table CSS identifier per workbook | multiple workbook HTML export Aspose.Cells | automated Excel HTML generation .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The sample iterates over a collection of Excel file paths, ensures the output directory exists, loads each workbook with Aspose.Cells, sets HtmlSaveOptions.TableCssId to a unique value (e.g., table1, table2), and saves each workbook as an individual HTML file. It logs success messages and handles missing files or other errors, enabling batch conversion of Excel workbooks to HTML with distinct table CSS IDs.
class BatchHtmlExport
{
    static void Main()
    {
        try
        {
            // List of source workbook file paths
            List<string> workbookPaths = new List<string>
            {
                @"C:\Workbooks\Report1.xlsx",
                @"C:\Workbooks\Report2.xlsx",
                @"C:\Workbooks\Report3.xlsx"
            };

            // Destination folder for HTML files
            string outputFolder = @"C:\HtmlExports\";

            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            // Counter to generate distinct TableCssId values
            int tableIdCounter = 1;

            foreach (string wbPath in workbookPaths)
            {
                try
                {
                    // Verify the workbook file exists
                    if (!File.Exists(wbPath))
                    {
                        Console.WriteLine($"File not found: {wbPath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(wbPath);

                    // Configure HTML save options
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
                    {
                        // Assign a unique TableCssId for each workbook
                        TableCssId = $"table{tableIdCounter}"
                        // By default Aspose.Cells exports HTML, CSS and images as separate files
                    };

                    // Build the output HTML file path
                    string htmlFileName = Path.GetFileNameWithoutExtension(wbPath) + ".html";
                    string htmlPath = Path.Combine(outputFolder, htmlFileName);

                    // Save the workbook as HTML
                    workbook.Save(htmlPath, saveOptions);

                    Console.WriteLine($"Exported '{wbPath}' to HTML with TableCssId='table{tableIdCounter}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{wbPath}': {ex.Message}");
                }

                tableIdCounter++;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
