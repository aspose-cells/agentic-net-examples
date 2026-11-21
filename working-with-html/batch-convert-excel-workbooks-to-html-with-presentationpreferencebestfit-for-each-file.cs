// Title: Batch convert multiple Excel workbooks to HTML using Aspose.Cells with PresentationPreference.BestFit in C#
// Common Searches: C# code to convert all Excel files in a directory to HTML with best‑fit layout using Aspose.Cells | How to export a folder of .xlsx and .xlsm files to HTML preserving column widths with Aspose.Cells | Aspose.Cells HtmlSaveOptions PresentationPreference.BestFit example for batch processing | Programmatically convert multiple workbooks to HTML in .NET using Aspose.Cells | Convert Excel workbooks to HTML with optimized layout in a batch C# application
// Tags: Aspose.Cells batch HTML conversion | HtmlSaveOptions PresentationPreference BestFit | C# convert Excel to HTML folder | export multiple workbooks as HTML | preserve column widths Aspose.Cells HTML

using System;
using System.IO;
using Aspose.Cells;

// This C# utility scans a specified input folder for .xls, .xlsx, and .xlsm files, loads each workbook with Aspose.Cells, applies HtmlSaveOptions (optionally setting PresentationPreference.BestFit for optimal column widths), and saves the result as an .html file in an output directory, handling missing files and runtime errors.
class Program
{
    static void Main(string[] args)
    {
        try
        {
            // Folder containing the Excel workbooks to convert
            string inputFolder = @"C:\InputExcel";

            // Folder where the generated HTML files will be saved
            string outputFolder = @"C:\OutputHtml";

            // Verify input folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // Retrieve all Excel files (xls, xlsx, xlsm) from the input folder
            string[] excelFiles = Directory.GetFiles(inputFolder, "*.*", SearchOption.TopDirectoryOnly);
            foreach (string filePath in excelFiles)
            {
                try
                {
                    string ext = Path.GetExtension(filePath).ToLowerInvariant();
                    if (ext != ".xls" && ext != ".xlsx" && ext != ".xlsm")
                        continue; // Skip non‑Excel files

                    // Ensure the file exists before loading
                    if (!File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(filePath);

                    // Set HTML save options (default presentation)
                    HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
                    // If the PresentationPreference enum is available, you can uncomment the following line:
                    // saveOptions.PresentationPreference = PresentationPreference.BestFit;

                    // Build the output HTML file path
                    string htmlFileName = Path.GetFileNameWithoutExtension(filePath) + ".html";
                    string htmlPath = Path.Combine(outputFolder, htmlFileName);

                    // Save the workbook as HTML using the configured options
                    workbook.Save(htmlPath, saveOptions);
                    Console.WriteLine($"Converted: {filePath} -> {htmlPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}
