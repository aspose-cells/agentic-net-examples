// Title: Open an Excel workbook from a ZIP archive with Aspose.Cells (C# stream example)
// Description: Demonstrates how to verify a ZIP file, locate a .xlsx entry with System.IO.Compression.ZipArchive, extract the entry as a stream, and load it directly into an Aspose.Cells Workbook. The sample reads the first worksheet name and cell A1, then optionally saves the workbook to a new file, all without writing the original Excel file to disk.
// Keywords: Aspose.Cells load workbook from zip stream | C# extract Excel from zip archive | Workbook constructor stream parameter | read Excel entry using ZipArchive | open workbook without extracting to disk | Aspose.Cells ZipArchive example | C# in‑memory Excel processing
// Common Searches: How to open an Excel file inside a zip with Aspose.Cells .NET | C# load workbook from zip entry stream | Aspose.Cells read .xlsx from ZipArchive | Open Excel from zip without extracting file | Aspose.Cells example for zip archive
// Developer Intent: Load an Excel workbook directly from a ZIP archive using a memory stream, avoiding temporary files.
// Use Cases: Read worksheet names or cell values from Excel files packaged in a zip archive. | Process multiple .xlsx entries in a zip, modify them in memory, and save the results. | Perform calculations on a workbook extracted from a zip and export the updated file without intermediate extraction.
// AI Prompts: Generate C# code that iterates over all .xlsx entries in a zip file and opens each with Aspose.Cells Workbook. | Provide robust error‑handling patterns for loading a workbook from a zip entry stream using Aspose.Cells. | Show how to change a cell value in the extracted workbook and save the updated file back to disk without extracting the original Excel file.

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

namespace AsposeCellsZipExample
{
    // Demonstrates how to verify a ZIP file, locate a .xlsx entry with System.IO.Compression.ZipArchive, extract the entry as a stream, and load it directly into an Aspose.Cells Workbook. The sample reads the first worksheet name and cell A1, then optionally saves the workbook to a new file, all without writing the original Excel file to disk.
    class Program
    {
        static void Main()
        {
            // Path to the ZIP archive that contains an Excel file
            string zipPath = "sample.zip";

            // Name of the Excel file inside the ZIP archive
            string excelEntryName = "sample.xlsx";

            try
            {
                // Verify that the ZIP file exists before attempting to open it
                if (!File.Exists(zipPath))
                {
                    Console.WriteLine($"ZIP archive not found: '{zipPath}'.");
                    return;
                }

                // Open the ZIP archive for reading
                using (FileStream zipFileStream = new FileStream(zipPath, FileMode.Open, FileAccess.Read))
                using (ZipArchive zipArchive = new ZipArchive(zipFileStream, ZipArchiveMode.Read))
                {
                    // Locate the specific entry (Excel file) within the archive
                    ZipArchiveEntry excelEntry = zipArchive.GetEntry(excelEntryName);
                    if (excelEntry == null)
                    {
                        Console.WriteLine($"Entry '{excelEntryName}' not found in the ZIP archive.");
                        return;
                    }

                    // Extract the entry as a stream
                    using (Stream entryStream = excelEntry.Open())
                    {
                        // Ensure the stream is positioned at the beginning
                        if (entryStream.CanSeek)
                            entryStream.Seek(0, SeekOrigin.Begin);

                        // Load the workbook from the extracted stream
                        Workbook workbook = new Workbook(entryStream);

                        // Access the first worksheet and read a cell value (for demonstration)
                        Worksheet sheet = workbook.Worksheets[0];
                        Console.WriteLine("First worksheet name: " + sheet.Name);
                        Console.WriteLine("Cell A1 value: " + sheet.Cells["A1"].StringValue);

                        // Optionally, save the workbook to a new file
                        string outputPath = "ExtractedWorkbook.xlsx";
                        workbook.Save(outputPath, SaveFormat.Xlsx);
                        Console.WriteLine($"Workbook saved to '{outputPath}'.");
                    }
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
