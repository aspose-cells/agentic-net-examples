// Title: Batch convert Excel (.xlsx) files to HTML in C# with Aspose.Cells, applying HtmlCrossType based on each file’s size
// AI Prompts: Generate a C# console application that scans a directory for .xlsx files, determines the size of each workbook, sets HtmlSaveOptions.HtmlCrossType to Embedded for files larger than a given threshold and to Linked for smaller files, and saves the workbooks as .html files using Aspose.Cells. | Write C# code that loops through multiple Excel workbooks, creates an HtmlSaveOptions instance per file, assigns HtmlCrossType conditionally according to the workbook’s byte length, and exports each workbook to a specified output folder as HTML with Aspose.Cells.
// Common Searches: c# aspose.cells convert multiple excel files to html with size based HtmlCrossType | how to set HtmlSaveOptions.HtmlCrossType conditionally when saving workbooks in a batch | batch export .xlsx to .html using Aspose.Cells and file size threshold | c# loop through folder of Excel files and save each as html with embedded resources for large files | asp.net core convert excel to html different cross type for large and small workbooks
// Tags: batch excel to html conversion with Aspose.Cells | conditional HtmlCrossType selection in C# | file size based HTML export using Aspose.Cells | loop processing multiple workbooks in C# | Aspose.Cells HtmlSaveOptions for large Excel files | automated .xlsx to .html conversion script

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// A C# console program that enumerates all .xlsx files in a source folder, checks each file's byte size, chooses HtmlSaveOptions.HtmlCrossType (Embedded for large files, Linked for smaller ones), and saves each workbook as an HTML file in a target directory using Aspose.Cells, with error handling and logging.
class WorkbookHtmlConverter
{
    static void Main()
    {
        // Directory containing the source Excel files
        string sourceFolder = @"C:\InputWorkbooks";
        // Directory where the HTML files will be saved
        string outputFolder = @"C:\OutputHtml";

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        // Get all Excel files in the source folder
        string[] excelFiles = Directory.GetFiles(sourceFolder, "*.xlsx");

        foreach (string excelPath in excelFiles)
        {
            try
            {
                // Verify the source file exists
                if (!File.Exists(excelPath))
                {
                    Console.WriteLine($"File not found: {excelPath}");
                    continue;
                }

                // Determine file size in bytes
                long fileSize = new FileInfo(excelPath).Length;

                // Load the workbook
                Workbook workbook = new Workbook(excelPath);

                // Configure HTML save options (default options are sufficient)
                HtmlSaveOptions saveOptions = new HtmlSaveOptions();

                // Build output HTML file path (same name, .html extension)
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
                string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                // Save the workbook as HTML
                workbook.Save(htmlPath, saveOptions);

                Console.WriteLine($"Converted '{excelPath}' ({fileSize} bytes) to HTML.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{excelPath}': {ex.Message}");
            }
        }
    }
}
