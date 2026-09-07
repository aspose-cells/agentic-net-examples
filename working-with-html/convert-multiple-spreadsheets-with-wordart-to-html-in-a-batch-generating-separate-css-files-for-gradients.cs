// Title: Batch convert Excel workbooks with WordArt to HTML and generate separate CSS files for gradients using Aspose.Cells for .NET
// AI Prompts: Write a C# console application that scans a folder for .xls, .xlsx, and .xlsm files, loads each workbook with Aspose.Cells, and saves it as HTML with HtmlSaveOptions configured to export images as files and create external CSS files for WordArt gradients. | Enhance the sample code to record the outcome of each conversion (success or error) in a log file, including source path, generated HTML path, and any exception details. | Add a command‑line switch that tells the batch converter to ignore hidden worksheets and export only visible sheets to HTML, using the appropriate Aspose.Cells settings.
// Common Searches: how to batch convert multiple Excel files to HTML with Aspose.Cells in C# | export WordArt gradient styles to separate CSS when saving Excel as HTML | Aspose.Cells HtmlSaveOptions disable base64 image embedding | C# program to process all .xlsx files in a folder and generate HTML and CSS output | skip hidden worksheets during Excel to HTML conversion using Aspose.Cells
// Tags: batch excel html conversion Aspose.Cells | external css for wordart gradients | disable base64 image export htmlsaveoptions | c# iterate workbook files in folder | skip hidden worksheets during html export

using System;
using System.IO;
using Aspose.Cells;

namespace BatchExcelToHtml
{
    // The example scans a given input directory for .xls, .xlsx, and .xlsm workbooks, loads each with Aspose.Cells, and saves them as HTML using HtmlSaveOptions that export images as separate files and write WordArt gradient styles to external CSS files. The resulting .html and .css files are placed in a designated output folder, enabling batch processing of spreadsheets with rich WordArt formatting.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder that contains the source Excel files
            string inputFolder = @"C:\InputExcels";

            // Folder where the HTML and CSS files will be written
            string outputFolder = @"C:\OutputHtml";

            try
            {
                // Ensure the output directory exists
                Directory.CreateDirectory(outputFolder);

                // Verify input folder exists
                if (!Directory.Exists(inputFolder))
                {
                    Console.WriteLine($"Input folder does not exist: {inputFolder}");
                    return;
                }

                // Get all Excel files in the input folder
                string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
                foreach (string excelPath in excelFiles)
                {
                    // Process only supported Excel extensions
                    string ext = Path.GetExtension(excelPath).ToLowerInvariant();
                    if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsm")
                        continue;

                    // Ensure the file actually exists before loading
                    if (!File.Exists(excelPath))
                    {
                        Console.WriteLine($"File not found: {excelPath}");
                        continue;
                    }

                    try
                    {
                        // Load the workbook
                        Workbook workbook = new Workbook(excelPath);

                        // Prepare HTML save options
                        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
                        {
                            // Export all worksheets
                            ExportActiveWorksheetOnly = false,

                            // Do not embed images as Base64; they will be saved as separate files
                            ExportImagesAsBase64 = false
                        };

                        // Determine output file names
                        string htmlFileName = Path.GetFileNameWithoutExtension(excelPath) + ".html";
                        string htmlOutputPath = Path.Combine(outputFolder, htmlFileName);

                        // Save the workbook as HTML (this will also create the CSS file)
                        workbook.Save(htmlOutputPath, saveOptions);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing file '{excelPath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
