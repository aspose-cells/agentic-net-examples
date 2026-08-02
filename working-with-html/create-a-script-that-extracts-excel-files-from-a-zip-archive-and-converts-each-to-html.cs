// Title: C# – Extract Excel files from a ZIP and batch‑convert to HTML with Aspose.Cells
// Description: A console program that validates a ZIP archive, extracts every workbook (.xlsx, .xls, .xlsm, .xlsb) to a temporary folder, and uses Aspose.Cells.Utility.ConversionUtility to convert each file to an HTML document. The script logs successful conversions, handles per‑file errors, and works without manual file handling.
// Keywords: Aspose.Cells C# | convert Excel to HTML | batch Excel HTML conversion | extract Excel from zip | .NET zip archive processing | ConversionUtility Convert method | automated spreadsheet to web preview
// Common Searches: C# extract Excel files from zip and convert to HTML | Aspose.Cells batch conversion from zip archive | how to convert multiple Excel workbooks to HTML in .NET | using ConversionUtility to generate HTML from Excel files | automate Excel to HTML conversion from compressed files
// Developer Intent: The developer needs a script that pulls all Excel workbooks out of a ZIP file and converts each one to an HTML file using Aspose.Cells, with minimal manual steps and robust error handling.
// Use Cases: Generate web‑ready previews of a large collection of spreadsheets delivered as a ZIP package. | Automate reporting pipelines where Excel reports are archived and must be published as HTML pages. | Create a temporary extraction workflow that processes each workbook, converts it to HTML, and logs any failures for later review.
// AI Prompts: Write C# code that extracts .xlsx, .xls, .xlsm, and .xlsb files from a ZIP archive and converts each to HTML using Aspose.Cells.Utility.ConversionUtility, including proper disposal of resources. | Refactor the script to use async/await for file I/O while preserving error handling and temporary folder cleanup. | Show how to modify the program to output all HTML files to a user‑specified directory while keeping the original workbook names.

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells.Utility;

// A console program that validates a ZIP archive, extracts every workbook (.xlsx, .xls, .xlsm, .xlsb) to a temporary folder, and uses Aspose.Cells.Utility.ConversionUtility to convert each file to an HTML document. The script logs successful conversions, handles per‑file errors, and works without manual file handling.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the zip archive containing Excel files
            string zipPath = "input.zip";

            // Verify that the zip file exists before proceeding
            if (!File.Exists(zipPath))
            {
                Console.WriteLine($"Error: The zip archive '{zipPath}' was not found.");
                return;
            }

            // Temporary folder to extract Excel files
            string extractDir = Path.Combine(Path.GetTempPath(), "ExcelExtract");
            Directory.CreateDirectory(extractDir);

            // Open the zip archive for reading
            using (ZipArchive archive = ZipFile.OpenRead(zipPath))
            {
                foreach (var entry in archive.Entries)
                {
                    // Process only Excel files
                    if (IsExcelFile(entry.FullName))
                    {
                        try
                        {
                            // Extract the Excel file to the temporary folder
                            string extractedPath = Path.Combine(extractDir, entry.Name);
                            entry.ExtractToFile(extractedPath, true);

                            // Determine the output HTML file path
                            string htmlPath = Path.ChangeExtension(extractedPath, ".html");

                            // Convert the Excel file to HTML using Aspose.Cells ConversionUtility
                            ConversionUtility.Convert(extractedPath, htmlPath);

                            Console.WriteLine($"Converted '{entry.Name}' to '{htmlPath}'.");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Failed to convert '{entry.Name}': {ex.Message}");
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    // Helper method to identify Excel file extensions
    static bool IsExcelFile(string fileName)
    {
        string ext = Path.GetExtension(fileName).ToLowerInvariant();
        return ext == ".xlsx" || ext == ".xls" || ext == ".xlsm" || ext == ".xlsb";
    }
}
