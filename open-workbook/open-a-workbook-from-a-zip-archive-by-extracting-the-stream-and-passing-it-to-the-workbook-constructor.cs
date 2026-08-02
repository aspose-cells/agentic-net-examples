// Title: Load an Excel workbook from a ZIP archive using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to open a .xlsx file stored inside a .zip package by reading the entry as a Stream with System.IO.Compression, passing the stream to the Aspose.Cells Workbook constructor, accessing cell values, and saving the workbook. Includes checks for missing files or entries and basic error handling.
// Keywords: Aspose.Cells | C# | .NET | ZipArchive | Workbook from stream | open Excel from zip | extract xlsx entry | read cell A1 | save workbook | error handling
// Common Searches: Aspose.Cells load Excel from zip archive | C# open .xlsx inside .zip without extracting | Workbook constructor stream Aspose.Cells example | read cell from zipped Excel file | extract Excel entry using ZipArchive in .NET
// Developer Intent: Read an Excel workbook directly from a ZIP file by extracting the .xlsx entry as a stream and initializing an Aspose.Cells Workbook with that stream.
// Use Cases: Read or modify data in Excel files that are distributed as part of a zip package without creating temporary files. | Batch‑process multiple .xlsx files inside a zip archive for conversion, validation, or data extraction. | Integrate zipped Excel resources into web services or APIs where disk I/O must be minimized.
// AI Prompts: Generate C# code to loop through all .xlsx entries in a zip archive and open each with Aspose.Cells. | Show how to open a password‑protected Excel file inside a zip using Aspose.Cells. | Provide best‑practice error handling for loading workbooks from ZipArchive streams in .NET.

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to open a .xlsx file stored inside a .zip package by reading the entry as a Stream with System.IO.Compression, passing the stream to the Aspose.Cells Workbook constructor, accessing cell values, and saving the workbook. Includes checks for missing files or entries and basic error handling.
    public class OpenWorkbookFromZip
    {
        public static void Run()
        {
            try
            {
                // Path to the ZIP archive containing the Excel file
                string zipFilePath = "sample.zip";

                // Name of the Excel file entry inside the ZIP archive
                string excelEntryName = "sample.xlsx";

                // Verify that the ZIP file exists
                if (!File.Exists(zipFilePath))
                {
                    Console.WriteLine($"ZIP file '{zipFilePath}' not found.");
                    return;
                }

                // Open the ZIP archive for reading
                using (FileStream zipStream = new FileStream(zipFilePath, FileMode.Open, FileAccess.Read))
                using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Read))
                {
                    // Locate the specific entry (Excel file) within the archive
                    ZipArchiveEntry excelEntry = archive.GetEntry(excelEntryName);
                    if (excelEntry == null)
                    {
                        Console.WriteLine($"Entry '{excelEntryName}' not found in the ZIP archive.");
                        return;
                    }

                    // Open a stream to the Excel entry
                    using (Stream excelStream = excelEntry.Open())
                    {
                        // Create a Workbook instance from the extracted stream
                        Workbook workbook = new Workbook(excelStream);

                        // Example operation: read the value of cell A1 from the first worksheet
                        Worksheet sheet = workbook.Worksheets[0];
                        Console.WriteLine($"Cell A1 value: {sheet.Cells["A1"].StringValue}");

                        // Save the workbook to a new file on disk
                        string outputPath = "ExtractedWorkbook.xlsx";
                        workbook.Save(outputPath);
                        Console.WriteLine($"Workbook saved to '{outputPath}'.");
                    }
                }

                Console.WriteLine("Workbook extracted from ZIP and saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            OpenWorkbookFromZip.Run();
        }
    }
}
