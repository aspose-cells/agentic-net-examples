// Title: C# – Batch Convert All XLSX Files in a Folder to HTML with Aspose.Cells ConversionUtility
// Description: A C# console program that scans a specified directory for *.xlsx files, creates an output folder when needed, and uses Aspose.Cells.Utility.ConversionUtility.Convert (default settings) to produce matching .html files while logging successes and errors.
// Keywords: Aspose.Cells batch conversion | C# convert xlsx to html | ConversionUtility.Convert example | bulk Excel to HTML | default HTML export Aspose.Cells | process multiple Excel files | directory conversion Aspose.Cells | C# console batch convert | Aspose.Cells HTML export | convert folder of Excel files
// Common Searches: batch convert xlsx to html c# | Aspose.Cells convert multiple Excel files to HTML | C# code to export all workbooks in a folder as HTML | how to use ConversionUtility.Convert for batch processing | Aspose.Cells default HTML export example
// Developer Intent: Create HTML versions of every Excel workbook in a given folder using Aspose.Cells default conversion settings.
// Use Cases: Publish a set of Excel‑based reports on a website without manual conversion. | Automate archival of spreadsheet data as web‑ready HTML snapshots. | Generate static HTML documentation from a collection of template workbooks stored on a server.
// AI Prompts: Show how to extend the sample to process subfolders recursively while preserving the folder hierarchy in the output. | Demonstrate adding custom CSS or a stylesheet link to each generated HTML file during batch conversion. | Explain techniques for monitoring progress and handling very large Excel files in a high‑volume batch conversion loop.

using System;
using System.IO;
using Aspose.Cells.Utility;

namespace AsposeCellsBatchConversion
{
    // A C# console program that scans a specified directory for *.xlsx files, creates an output folder when needed, and uses Aspose.Cells.Utility.ConversionUtility.Convert (default settings) to produce matching .html files while logging successes and errors.
    class Program
    {
        static void Main(string[] args)
        {
            // Directory containing the XLSX files.
            // Change this path as needed.
            string sourceDirectory = @"C:\InputXlsxFiles";

            // Directory where the HTML files will be saved.
            // It will be created automatically if it does not exist.
            string outputDirectory = @"C:\OutputHtmlFiles";

            // Verify source directory exists.
            if (!Directory.Exists(sourceDirectory))
            {
                Console.WriteLine($"Source directory '{sourceDirectory}' does not exist.");
                return;
            }

            // Ensure the output directory exists.
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            try
            {
                // Get all .xlsx files in the source directory (non‑recursive).
                string[] xlsxFiles = Directory.GetFiles(sourceDirectory, "*.xlsx", SearchOption.TopDirectoryOnly);

                foreach (string xlsxPath in xlsxFiles)
                {
                    // Determine the output HTML file name.
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(xlsxPath);
                    string htmlPath = Path.Combine(outputDirectory, fileNameWithoutExt + ".html");

                    try
                    {
                        // Convert the XLSX file to HTML using default conversion settings.
                        // The ConversionUtility.Convert method handles loading and saving internally.
                        ConversionUtility.Convert(xlsxPath, htmlPath);
                        Console.WriteLine($"Converted '{xlsxPath}' to '{htmlPath}'.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to convert '{xlsxPath}': {ex.Message}");
                    }
                }

                Console.WriteLine("Batch conversion completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during batch processing: {ex.Message}");
            }
        }
    }
}
