using System;
using System.IO;
using Aspose.Cells;

class BatchWordArtToHtml
{
    static void Main()
    {
        // Folder containing source Excel files (with WordArt)
        string sourceFolder = @"C:\InputSpreadsheets";

        // Folder where HTML files and associated CSS/images will be saved
        string outputFolder = @"C:\HtmlOutput";

        // Ensure the output folder exists
        Directory.CreateDirectory(outputFolder);

        // Verify source folder exists
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder not found: {sourceFolder}");
            return;
        }

        // Process each Excel file in the source folder
        foreach (string excelPath in Directory.GetFiles(sourceFolder, "*.*", SearchOption.TopDirectoryOnly))
        {
            // Consider only supported Excel formats
            string ext = Path.GetExtension(excelPath).ToLowerInvariant();
            if (ext != ".xlsx" && ext != ".xlsm" && ext != ".xls" && ext != ".xlsb")
                continue;

            // Ensure the file exists before loading
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"File not found: {excelPath}");
                continue;
            }

            try
            {
                // Load the workbook
                using (Workbook workbook = new Workbook(excelPath))
                {
                    // Prepare HTML save options
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        // Export each worksheet's CSS to a separate file (required for gradients)
                        ExportWorksheetCSSSeparately = true,
                        // Keep CSS in external files (not a single HTML file)
                        SaveAsSingleFile = false
                    };

                    // Store images and other attached files in a dedicated sub‑folder per workbook
                    string attachedDir = Path.Combine(outputFolder,
                        Path.GetFileNameWithoutExtension(excelPath) + "_files");
                    Directory.CreateDirectory(attachedDir);
                    htmlOptions.AttachedFilesDirectory = attachedDir;
                    htmlOptions.AttachedFilesUrlPrefix = Path.GetFileName(attachedDir) + "/";

                    // Define the HTML file name
                    string htmlFileName = Path.Combine(outputFolder,
                        Path.GetFileNameWithoutExtension(excelPath) + ".html");

                    // Save the workbook as HTML
                    workbook.Save(htmlFileName, htmlOptions);
                }

                Console.WriteLine($"Converted '{excelPath}' to HTML with separate CSS.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{excelPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}