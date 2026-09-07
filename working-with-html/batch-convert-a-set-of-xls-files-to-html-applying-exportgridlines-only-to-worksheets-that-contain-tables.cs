// Title: Batch convert XLS workbooks to HTML with grid lines only on worksheets that contain tables using Aspose.Cells for .NET
// AI Prompts: Generate C# code that scans a folder for .xls files, loads each workbook with Aspose.Cells, sets IsGridlinesVisible on worksheets that have ListObjects, enables ExportGridLines, and saves each workbook as an HTML file. | Create a C# utility that performs bulk XLS‑to‑HTML conversion, applying grid‑line visibility conditionally based on the presence of tables in each worksheet.
// Common Searches: how to export grid lines only for worksheets that contain tables using Aspose.Cells C# | batch convert multiple .xls files to html with conditional gridlines in .NET | detect ListObjects in a worksheet before saving as HTML with Aspose.Cells | C# Aspose.Cells hide gridlines on sheets without tables during HTML export | convert folder of Excel .xls to HTML preserving table formatting Aspose.Cells
// Tags: batch xls to html conversion Aspose.Cells | conditional gridlines per worksheet Aspose.Cells | detect listobjects worksheet Aspose.Cells | htmlsaveoptions exportgridlines true | isgridlinesvisible per sheet C#

using System;
using System.IO;
using Aspose.Cells;

// The program iterates over all .xls files in a specified input directory, loads each workbook with Aspose.Cells, toggles IsGridlinesVisible on worksheets that contain ListObjects, configures HtmlSaveOptions to export grid lines, and saves each workbook as an .html file in an output folder, handling missing files and exceptions.
class BatchXlsToHtmlConverter
{
    static void Main()
    {
        // Folder containing the source XLS files
        string inputFolder = @"C:\InputXls";
        // Folder where the HTML files will be saved
        string outputFolder = @"C:\OutputHtml";

        // Verify input folder exists
        if (!Directory.Exists(inputFolder))
        {
            Console.WriteLine($"Input folder not found: {inputFolder}");
            return;
        }

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Get all XLS files in the input folder
        string[] xlsFiles = Directory.GetFiles(inputFolder, "*.xls");

        foreach (string xlsPath in xlsFiles)
        {
            // Verify the file still exists before processing
            if (!File.Exists(xlsPath))
            {
                Console.WriteLine($"File not found (skipped): {xlsPath}");
                continue;
            }

            try
            {
                // Load the workbook from the XLS file
                Workbook workbook = new Workbook(xlsPath);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine if the worksheet contains any tables (ListObjects)
                    bool containsTable = sheet.ListObjects.Count > 0;

                    // Apply grid line visibility only when a table exists
                    // ExportGridLines is a global option, so we control per‑sheet visibility via IsGridlinesVisible
                    sheet.IsGridlinesVisible = containsTable;
                }

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Enable grid line export globally; per‑sheet visibility is handled above
                    ExportGridLines = true
                };

                // Build the output HTML file path
                string htmlFileName = Path.GetFileNameWithoutExtension(xlsPath) + ".html";
                string htmlPath = Path.Combine(outputFolder, htmlFileName);

                // Save the workbook as HTML
                workbook.Save(htmlPath, htmlOptions);
                Console.WriteLine($"Converted: {Path.GetFileName(xlsPath)} -> {htmlFileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file '{xlsPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
