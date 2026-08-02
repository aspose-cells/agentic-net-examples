// Title: Batch Convert Excel Files with WordArt to HTML (Separate CSS for Gradients) – C# Aspose.Cells
// Description: A C# console utility that scans a folder for .xlsx, .xlsm and .xls workbooks, loads each with Aspose.Cells, and saves them as HTML pages. The HtmlSaveOptions are set to export each worksheet’s CSS to its own file and enable CSS custom properties, so gradient styles from WordArt are stored in separate CSS files alongside the HTML output.
// Keywords: Aspose.Cells batch conversion | Excel to HTML C# | WordArt HTML export | ExportWorksheetCSSSeparately | EnableCssCustomProperties | gradient CSS Excel | HTML5 save options | automated Excel reporting | C# file system processing
// Common Searches: convert multiple Excel files with WordArt to HTML using Aspose.Cells | Aspose.Cells HtmlSaveOptions separate CSS files for gradients | C# batch Excel to HTML conversion script | export WordArt styles to CSS when saving Excel as HTML | how to generate CSS folders for each worksheet in Aspose.Cells
// Developer Intent: Automatically transform every Excel workbook in a directory—including those containing WordArt—into individual HTML pages with dedicated CSS files for gradient definitions.
// Use Cases: Mass‑publish Excel‑based dashboards to a web portal while preserving WordArt styling. | Integrate into a nightly build pipeline that converts financial reports to HTML for client‑side viewing. | Create a lightweight HTML archive of legacy spreadsheets, keeping gradient CSS separate for easier maintenance. | Skip non‑Excel files and log any load or save errors for later review.
// AI Prompts: Generate C# code that uses Aspose.Cells to batch convert Excel workbooks with WordArt to HTML, exporting each worksheet’s CSS to a separate file and enabling CSS custom properties for gradients. | Explain the impact of ExportWorksheetCSSSeparately and EnableCssCustomProperties on the HTML and CSS output when converting Excel to HTML with Aspose.Cells. | Provide a troubleshooting checklist for common failures (missing files, unsupported WordArt features, permission issues) during batch Excel‑to‑HTML conversion.

using System;
using System.IO;
using Aspose.Cells;

namespace BatchConversion
{
    // A C# console utility that scans a folder for .xlsx, .xlsm and .xls workbooks, loads each with Aspose.Cells, and saves them as HTML pages. The HtmlSaveOptions are set to export each worksheet’s CSS to its own file and enable CSS custom properties, so gradient styles from WordArt are stored in separate CSS files alongside the HTML output.
    class BatchWordArtToHtml
    {
        static void Main()
        {
            // Folder containing the source Excel files
            string inputFolder = @"C:\InputExcels";

            // Folder where HTML files and their CSS folders will be saved
            string outputFolder = @"C:\OutputHtml";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder not found: {inputFolder}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Retrieve all Excel files (xlsx, xlsm, xls) from the input folder
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string sourcePath in excelFiles)
            {
                string ext = Path.GetExtension(sourcePath).ToLowerInvariant();
                if (ext != ".xlsx" && ext != ".xlsm" && ext != ".xls")
                    continue; // Skip non‑Excel files

                // Verify the file still exists before loading
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"File not found (skipped): {sourcePath}");
                    continue;
                }

                try
                {
                    // Load the workbook
                    Workbook workbook = new Workbook(sourcePath);

                    // Configure HTML save options
                    HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                    {
                        // Export each worksheet's CSS to a separate file (gradients are stored in CSS)
                        ExportWorksheetCSSSeparately = true,

                        // Enable CSS custom properties to reuse gradient definitions efficiently
                        EnableCssCustomProperties = true,

                        // Use HTML5 for modern markup (optional)
                        HtmlVersion = HtmlVersion.Html5
                    };

                    // Determine the output HTML file path
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
                    string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                    // Save the workbook as HTML; Aspose creates a companion folder for CSS and images
                    workbook.Save(htmlPath, htmlOptions);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{sourcePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion of spreadsheets with WordArt to HTML completed.");
        }
    }
}
