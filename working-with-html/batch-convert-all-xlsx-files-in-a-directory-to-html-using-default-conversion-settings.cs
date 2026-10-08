// Title: Convert all XLSX files in a folder to HTML using Aspose.Cells for .NET (C# batch conversion)
// AI Prompts: Write a C# console program that enumerates every .xlsx file in a given directory, loads each workbook with Aspose.Cells, and saves it as an .html file using the default SaveFormat.Html, creating the output folder if needed. | Generate C# code that performs bulk Excel‑to‑HTML conversion with Aspose.Cells, handling missing source folders, creating destination folders, and logging success or error messages for each file.
// Common Searches: c# Aspose.Cells batch convert xlsx files in a directory to html | how to programmatically convert multiple Excel workbooks to HTML with Aspose.Cells | sample code for converting all .xlsx files in a folder to .html using Aspose.Cells .NET | Aspose.Cells saveformat.html bulk conversion example c#
// Tags: Aspose.Cells bulk xlsx to html conversion | C# directory enumeration for Excel files | Workbook.Save with SaveFormat.Html default settings | automatic output folder creation Aspose.Cells | per‑file error handling during batch conversion

using System;
using System.IO;
using Aspose.Cells;

// A C# console application that scans a specified input folder for .xlsx files, loads each workbook with Aspose.Cells, and saves it as an .html file in an output folder using the default HTML save format, with folder existence checks and per‑file error handling.
class BatchXlsxToHtmlConverter
{
    static void Main(string[] args)
    {
        // Folder containing the XLSX files. Adjust as needed.
        string sourceFolder = @"C:\InputXlsxFiles";

        // Folder where the HTML files will be saved.
        string outputFolder = @"C:\OutputHtmlFiles";

        // Verify source folder exists.
        if (!Directory.Exists(sourceFolder))
        {
            Console.WriteLine($"Source folder not found: {sourceFolder}");
            return;
        }

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        string[] xlsxFiles;
        try
        {
            // Get all .xlsx files in the source directory (non‑recursive).
            xlsxFiles = Directory.GetFiles(sourceFolder, "*.xlsx", SearchOption.TopDirectoryOnly);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error accessing source folder: {ex.Message}");
            return;
        }

        foreach (string xlsxPath in xlsxFiles)
        {
            try
            {
                // Load the XLSX workbook.
                Workbook workbook = new Workbook(xlsxPath);

                // Determine the output HTML file name.
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                string htmlPath = Path.Combine(outputFolder, fileNameWithoutExt + ".html");

                // Save the workbook as HTML.
                workbook.Save(htmlPath, SaveFormat.Html);

                Console.WriteLine($"Converted: {Path.GetFileName(xlsxPath)} -> {Path.GetFileName(htmlPath)}");
            }
            catch (Exception ex)
            {
                // Log any errors but continue processing remaining files.
                Console.WriteLine($"Error converting '{xlsxPath}': {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
