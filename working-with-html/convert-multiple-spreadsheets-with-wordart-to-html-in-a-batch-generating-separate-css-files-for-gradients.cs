// Title: Batch convert Excel files with WordArt to HTML5 using Aspose.Cells – separate CSS per worksheet
// Description: C# utility that scans a folder for .xlsx workbooks, loads each with Aspose.Cells, and saves them as HTML5 pages. The HtmlSaveOptions are set to export each worksheet’s CSS to its own file (preserving gradient styles), write WordArt and other images as external files, and omit unused styles for a lightweight output.
// Keywords: Aspose.Cells batch HTML export | C# Excel to HTML5 conversion | WordArt to HTML Aspose | ExportWorksheetCSSSeparately example | separate CSS per worksheet | external image export Aspose.Cells | ExcludeUnusedStyles usage | convert multiple Excel files | .NET Excel HTML conversion | gradient styles CSS Excel
// Common Searches: batch convert Excel to HTML with Aspose.Cells | export WordArt as external image C# | Aspose.Cells ExportWorksheetCSSSeparately tutorial | save Excel worksheets to separate CSS files | HTML5 output from multiple .xlsx files | Aspose.Cells exclude unused styles example | C# script to convert folder of Excel files to HTML
// Developer Intent: Automatically transform every Excel workbook in a directory into an HTML5 file, generating one CSS file per worksheet and saving WordArt and other graphics as separate image files.
// Use Cases: Publish a collection of Excel dashboards on a website while keeping each sheet’s gradient styling isolated in its own CSS file. | Schedule nightly batch exports of reporting workbooks to HTML for fast, cache‑friendly web delivery. | Create static documentation from Excel templates where WordArt must remain as high‑quality images rather than Base64 strings.
// AI Prompts: Generate C# code that uses Aspose.Cells to batch convert all .xlsx files in a directory to HTML5, exporting WordArt as external images and creating a distinct CSS file for each worksheet. | Explain how HtmlSaveOptions properties ExportWorksheetCSSSeparately, ExportImagesAsBase64, and ExcludeUnusedStyles affect the HTML output of workbooks containing WordArt. | Provide a step‑by‑step guide for configuring Aspose.Cells to produce separate CSS files for gradients when converting multiple Excel files to HTML.

using System;
using System.IO;
using Aspose.Cells;

// C# utility that scans a folder for .xlsx workbooks, loads each with Aspose.Cells, and saves them as HTML5 pages. The HtmlSaveOptions are set to export each worksheet’s CSS to its own file (preserving gradient styles), write WordArt and other images as external files, and omit unused styles for a lightweight output.
class BatchWordArtToHtml
{
    static void Main()
    {
        // Folder containing source Excel files with WordArt
        string sourceFolder = @"C:\InputExcels";

        // Folder where HTML output and separate CSS files will be saved
        string outputFolder = @"C:\OutputHtml";

        try
        {
            // Ensure the source folder exists; if not, inform the user and exit
            if (!Directory.Exists(sourceFolder))
            {
                Console.WriteLine($"Source folder not found: {sourceFolder}");
                return;
            }

            // Ensure the output folder exists
            Directory.CreateDirectory(outputFolder);

            // Process each Excel file in the source folder
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
                    // Load the workbook
                    using (Workbook workbook = new Workbook(excelPath))
                    {
                        // Configure HTML save options
                        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                        {
                            // Export each worksheet's CSS to a separate file (gradients will be in these CSS files)
                            ExportWorksheetCSSSeparately = true,

                            // Export images (including WordArt) as separate files instead of Base64 strings
                            ExportImagesAsBase64 = false,

                            // Use HTML5 for better standards compliance
                            HtmlVersion = HtmlVersion.Html5,

                            // Optional: keep unused styles excluded to reduce size
                            ExcludeUnusedStyles = true
                        };

                        // Determine output HTML file name (same as source file name with .html extension)
                        string htmlFileName = Path.GetFileNameWithoutExtension(excelPath) + ".html";
                        string htmlPath = Path.Combine(outputFolder, htmlFileName);

                        // Save the workbook as HTML using the configured options
                        workbook.Save(htmlPath, htmlOptions);

                        Console.WriteLine($"Converted '{excelPath}' to HTML at '{htmlPath}'.");
                    }
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
