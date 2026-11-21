// Title: C# batch conversion of all XLSX workbooks in a folder to HTML using Aspose.Cells default settings
// Common Searches: C# code to convert every .xlsx file in a directory to HTML with Aspose.Cells | How to perform batch XLSX to HTML conversion in .NET using Aspose.Cells | Automate conversion of multiple Excel workbooks to HTML files in a folder | Aspose.Cells example for processing a folder of Excel files and saving as HTML | Command‑line tool to convert a batch of XLSX files to HTML in C#
// Tags: Aspose.Cells batch XLSX to HTML conversion | C# directory enumeration for Excel files | Workbook.Save with SaveFormat.Html | default HTML export settings Aspose.Cells | automated Excel to HTML conversion .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample scans a specified input folder for .xlsx files, loads each workbook with Aspose.Cells, and saves it as an .html file in an output folder using the default Html SaveFormat. It creates missing directories, logs conversion progress, and continues processing despite individual file errors.
class BatchXlsxToHtmlConverter
{
    static void Main(string[] args)
    {
        // Source directory containing XLSX files (default or from command line).
        string sourceDirectory = args.Length > 0 ? args[0] : @"C:\InputXlsx";

        // Destination directory for generated HTML files (default or from command line).
        string outputDirectory = args.Length > 1 ? args[1] : @"C:\OutputHtml";

        try
        {
            // Ensure the source directory exists; if not, create it and inform the user.
            if (!Directory.Exists(sourceDirectory))
            {
                Directory.CreateDirectory(sourceDirectory);
                Console.WriteLine($"Source directory created: {sourceDirectory}");
                Console.WriteLine("Place XLSX files in the source directory and rerun the program.");
                return;
            }

            // Ensure the output directory exists.
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            // Retrieve all .xlsx files (non‑recursive) from the source directory.
            string[] xlsxFiles = Directory.GetFiles(sourceDirectory, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string xlsxPath in xlsxFiles)
            {
                try
                {
                    // Verify the file exists before attempting to load.
                    if (!File.Exists(xlsxPath))
                    {
                        Console.WriteLine($"File not found: {xlsxPath}");
                        continue;
                    }

                    // Load the workbook.
                    Workbook workbook = new Workbook(xlsxPath);

                    // Build the output HTML file path.
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                    string htmlPath = Path.Combine(outputDirectory, fileNameWithoutExt + ".html");

                    // Save as HTML using default settings.
                    workbook.Save(htmlPath, SaveFormat.Html);

                    Console.WriteLine($"Converted: {Path.GetFileName(xlsxPath)} -> {Path.GetFileName(htmlPath)}");
                }
                catch (Exception ex)
                {
                    // Log conversion errors but continue with other files.
                    Console.WriteLine($"Error converting '{xlsxPath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors (e.g., permission issues).
            Console.WriteLine($"Fatal error: {ex.Message}");
        }
    }
}
