// Title: Batch convert every XLSX workbook in a folder to HTML using Aspose.Cells for .NET (C#)
// AI Prompts: Write a C# console program that scans a given folder for *.xlsx files, loads each workbook with Aspose.Cells, and saves it as an .html file using the default SaveFormat.Html. | Create code that enumerates Excel files in a directory, converts each to HTML with Aspose.Cells, and writes the output to a separate folder while handling missing files and exceptions.
// Common Searches: asp.net batch conversion of xlsx files to html with aspose.cells | c# code to convert all Excel workbooks in a directory to html | how to export multiple .xlsx files to html using Aspose.Cells library | convert folder of Excel files to html programmatically in C# | default HTML export settings Aspose.Cells for batch processing
// Tags: Aspose.Cells XLSX to HTML batch conversion | C# directory enumeration for Excel to HTML export | Save workbook as HTML using Aspose.Cells default settings | Error handling for Excel to HTML conversion in C# | Create output folder for HTML files with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program enumerates all .xlsx files in a specified input folder, loads each workbook with Aspose.Cells, and saves it as an .html file in an output folder using the default HTML save format, with basic error handling.
class Program
{
    static void Main(string[] args)
    {
        // Folder containing the XLSX files. Change as needed or pass via args.
        string sourceFolder = @"C:\InputXlsx";
        // Folder where the HTML files will be saved.
        string outputFolder = @"C:\OutputHtml";

        // Verify source folder exists.
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder not found: {sourceFolder}");
            return;
        }

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Get all .xlsx files in the source folder (non‑recursive).
        string[] xlsxFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string xlsxPath in xlsxFiles)
        {
            // Verify the file still exists before loading.
            if (!File.Exists(xlsxPath))
            {
                Console.WriteLine($"File not found (skipped): {xlsxPath}");
                continue;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(xlsxPath);

                // Build the output HTML file name.
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                // Save the workbook as HTML.
                workbook.Save(htmlPath, SaveFormat.Html);

                Console.WriteLine($"Converted: {Path.GetFileName(xlsxPath)} -> {Path.GetFileName(htmlPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error converting '{xlsxPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
