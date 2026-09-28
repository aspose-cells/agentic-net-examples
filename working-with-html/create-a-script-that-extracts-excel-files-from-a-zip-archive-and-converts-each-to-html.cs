// Title: C# script to unzip Excel files and convert each workbook to HTML with Aspose.Cells
// AI Prompts: Create a C# console application that opens a ZIP archive, extracts every .xls or .xlsx file, loads each into an Aspose.Cells Workbook, and saves the workbook as an individual HTML file. | Add detailed logging to the zip‑to‑HTML conversion program to record which workbooks were successfully converted and which were skipped, while ignoring non‑Excel entries. | Modify the script to also handle .xlsm files and accept the output HTML directory as a command‑line parameter.
// Common Searches: how to batch convert Excel files inside a zip archive to HTML using Aspose.Cells C# | C# unzip multiple .xlsx files and save each as HTML with Aspose.Cells SaveFormat.Html | process compressed Excel workbooks and generate HTML reports in .NET
// Tags: Aspose.Cells unzip Excel to HTML conversion | C# extract .xlsx from ZIP archive | SaveFormat.Html workbook export | batch Excel to HTML with Aspose.Cells | skip non‑Excel entries during ZIP processing

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

// The example opens a specified ZIP file, iterates through its entries, loads each .xls or .xlsx workbook into an Aspose.Cells Workbook, and saves the workbook as an HTML file in a target folder, creating the folder if necessary and handling errors gracefully.
class ExcelZipToHtmlConverter
{
    static void Main(string[] args)
    {
        // Path to the zip archive containing Excel files
        string zipPath = @"C:\Path\To\Your\Archive.zip";

        // Directory where the generated HTML files will be saved
        string outputDirectory = @"C:\Path\To\Output\Html";

        try
        {
            // Verify that the zip archive exists
            if (!File.Exists(zipPath))
            {
                Console.WriteLine($"Error: Zip archive not found at '{zipPath}'.");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputDirectory);

            // Open the zip archive for reading
            using (ZipArchive archive = ZipFile.OpenRead(zipPath))
            {
                // Iterate through each entry in the zip archive
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    // Process only files with .xls or .xlsx extensions
                    string extension = Path.GetExtension(entry.FullName);
                    if (extension.Equals(".xls", StringComparison.OrdinalIgnoreCase) ||
                        extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        // Open the entry stream
                        using (Stream entryStream = entry.Open())
                        {
                            // Load the Excel file into an Aspose.Cells Workbook
                            Workbook workbook = new Workbook(entryStream);

                            // Determine the output HTML file path
                            string htmlFileName = Path.GetFileNameWithoutExtension(entry.Name) + ".html";
                            string htmlFilePath = Path.Combine(outputDirectory, htmlFileName);

                            // Save the workbook as an HTML file
                            workbook.Save(htmlFilePath, SaveFormat.Html);
                        }
                    }
                }
            }

            Console.WriteLine("Conversion completed. HTML files are located at: " + outputDirectory);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred during conversion:");
            Console.WriteLine(ex.Message);
        }
    }
}
