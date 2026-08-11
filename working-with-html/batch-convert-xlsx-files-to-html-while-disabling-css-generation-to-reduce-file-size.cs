// Title: Batch convert XLSX to HTML with CSS disabled using Aspose.Cells for .NET
// Description: A C# utility that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, and saves it as an HTML file using HtmlSaveOptions.DisableCss (inline styles only) and ExcludeUnusedStyles to minimize output size. The program creates the target directory if needed and logs conversion results and errors.
// Keywords: Aspose.Cells batch conversion | XLSX to HTML C# | DisableCss Aspose.Cells | HtmlSaveOptions inline styles | reduce HTML size Excel | convert multiple Excel files to HTML | Aspose.Cells .NET example
// Common Searches: convert all xlsx files in a folder to html asp.net | aspocells disable css html export batch | c# batch excel to html without external css | aspocells htmlsaveoptions disablecss example | how to shrink html output from aspocells conversion
// Developer Intent: Convert every XLSX workbook in a specified directory to an HTML file with only inline styling, eliminating external CSS to keep the output lightweight.
// Use Cases: Generate compact HTML previews of uploaded Excel files for web portals. | Create email‑ready HTML reports from a batch of Excel templates without external style sheets. | Automate server‑side conversion of Excel workbooks to inline‑styled HTML for PDF or document generation.
// AI Prompts: Write C# code that uses Aspose.Cells to convert all .xlsx files in a folder to .html with inline styles only, handling missing files and I/O errors. | Explain the impact of HtmlSaveOptions.DisableCss and ExcludeUnusedStyles on the size of generated HTML and how to verify they are applied. | Modify the batch conversion to process subfolders recursively while still disabling CSS generation.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchHtmlConversion
{
    // A C# utility that scans a folder for .xlsx files, loads each workbook with Aspose.Cells, and saves it as an HTML file using HtmlSaveOptions.DisableCss (inline styles only) and ExcludeUnusedStyles to minimize output size. The program creates the target directory if needed and logs conversion results and errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Example usage:
            // Source folder containing XLSX files
            string sourceFolder = @"C:\InputXlsx";
            // Destination folder for generated HTML files
            string outputFolder = @"C:\OutputHtml";

            BatchConvertXlsxToHtml(sourceFolder, outputFolder);
        }

        /// <param name="sourceFolder">Folder containing the source XLSX files.</param>
        /// <param name="outputFolder">Folder where the HTML files will be saved.</param>
        static void BatchConvertXlsxToHtml(string sourceFolder, string outputFolder)
        {
            // Verify source folder exists
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder does not exist: {sourceFolder}");
                return;
            }

            // Ensure the output directory exists
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            // Get all .xlsx files in the source directory (non‑recursive)
            string[] xlsxFiles;
            try
            {
                xlsxFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to enumerate files in '{sourceFolder}': {ex.Message}");
                return;
            }

            foreach (string sourcePath in xlsxFiles)
            {
                try
                {
                    // Verify the source file exists (defensive check)
                    if (!File.Exists(sourcePath))
                    {
                        Console.WriteLine($"File not found: {sourcePath}");
                        continue;
                    }

                    // Determine the output HTML file name (same base name, .html extension)
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                    string destPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                    // Load options for reading the XLSX file
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

                    // HTML save options with CSS disabled
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions
                    {
                        DisableCss = true,               // Use only inline styles
                        ExcludeUnusedStyles = true       // Exclude unused CSS (default true)
                    };

                    // Perform the conversion using the overload that accepts load and save options
                    ConversionUtility.Convert(sourcePath, loadOptions, destPath, saveOptions);

                    Console.WriteLine($"Converted: {sourcePath} -> {destPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error converting '{sourcePath}': {ex.Message}");
                }
            }
        }
    }
}
