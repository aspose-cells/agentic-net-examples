// Title: Load an .xlsx workbook directly from a ZIP archive stream using Aspose.Cells in C#
// AI Prompts: Write a C# method that receives a zip file path and an entry name, opens the entry as a stream, and returns an Aspose.Cells Workbook object. | Create a robust C# example that locates an .xlsx file inside a ZIP archive, loads it into a Workbook via Aspose.Cells, and saves it to a target location. | Generate C# code with proper exception handling for opening a workbook from a zip entry stream using Aspose.Cells.
// Common Searches: c# load xlsx from zip archive using aspose.cells | how to open an Excel file inside a zip without extracting to disk in .net | asp.net read workbook from zip entry stream | example of extracting .xlsx from zip and loading into Aspose.Cells | c# error handling when loading workbook from zip entry
// Tags: Aspose.Cells load workbook from zip stream | C# extract xlsx entry from zip archive | open Excel workbook from compressed file Aspose.Cells | read zip entry as stream for Aspose.Cells | save extracted workbook with Aspose.Cells

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

namespace MyApp
{
    // The sample verifies the ZIP file exists, opens it, finds the specified .xlsx entry, extracts the entry as a stream, creates an Aspose.Cells Workbook from that stream, and saves the workbook to a new .xlsx file, creating the output directory if necessary and handling any exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the ZIP archive containing the Excel file
                string zipPath = @"C:\Data\workbook.zip";

                // Name of the Excel file inside the ZIP archive
                string excelEntryName = "sample.xlsx";

                // Verify that the ZIP file exists
                if (!File.Exists(zipPath))
                {
                    Console.WriteLine($"ZIP file not found: {zipPath}");
                    return;
                }

                // Open the ZIP archive for reading
                using (FileStream zipFileStream = new FileStream(zipPath, FileMode.Open, FileAccess.Read))
                using (ZipArchive archive = new ZipArchive(zipFileStream, ZipArchiveMode.Read))
                {
                    // Locate the entry that holds the Excel workbook
                    ZipArchiveEntry excelEntry = archive.GetEntry(excelEntryName);
                    if (excelEntry == null)
                    {
                        Console.WriteLine($"Entry '{excelEntryName}' not found in the ZIP archive.");
                        return;
                    }

                    // Extract the entry as a stream and load it into Aspose.Cells Workbook
                    using (Stream excelStream = excelEntry.Open())
                    {
                        Workbook workbook = new Workbook(excelStream);

                        // Save the extracted workbook to a new file
                        string outputPath = @"C:\Data\extracted_sample.xlsx";
                        string outputDir = Path.GetDirectoryName(outputPath);
                        if (!Directory.Exists(outputDir))
                        {
                            Directory.CreateDirectory(outputDir);
                        }

                        workbook.Save(outputPath);
                        Console.WriteLine($"Workbook extracted and saved to '{outputPath}'.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
