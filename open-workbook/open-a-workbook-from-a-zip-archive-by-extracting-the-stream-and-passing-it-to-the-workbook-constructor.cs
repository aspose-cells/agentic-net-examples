// Title: Open an Excel workbook stored inside a ZIP file by extracting the entry stream and initializing Aspose.Cells Workbook in C#
// Common Searches: C# Aspose.Cells load Excel file from zip archive stream | How to read a workbook from a zip entry using Aspose.Cells .NET | Extract .xlsx from ZipArchive and open with Aspose.Cells | Aspose.Cells Workbook constructor with Stream parameter in C# | Open .xlsx inside .zip without extracting to disk using Aspose.Cells
// Tags: Aspose.Cells load workbook from ZipArchive stream | initialize Workbook with Stream in C# | extract .xlsx entry from zip using System.IO.Compression | open compressed Excel file without extracting to disk | read first worksheet name after loading from zip

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

// The example verifies the ZIP file, opens it with System.IO.Compression.ZipArchive, locates the specified .xlsx entry, obtains its data as a Stream, creates an Aspose.Cells Workbook from that Stream, and demonstrates accessing the first worksheet name (with optional saving).
class Program
{
    static void Main()
    {
        try
        {
            // Path to the ZIP archive containing the Excel file
            string zipPath = @"C:\Data\WorkbookArchive.zip";

            // Verify that the ZIP file exists
            if (!File.Exists(zipPath))
            {
                Console.WriteLine($"ZIP file not found: {zipPath}");
                return;
            }

            // Name of the Excel file inside the ZIP archive
            string excelEntryName = "MyWorkbook.xlsx";

            // Open the ZIP archive for reading
            using (FileStream zipFileStream = new FileStream(zipPath, FileMode.Open, FileAccess.Read))
            using (ZipArchive archive = new ZipArchive(zipFileStream, ZipArchiveMode.Read))
            {
                // Locate the specific entry that holds the workbook
                ZipArchiveEntry excelEntry = archive.GetEntry(excelEntryName);
                if (excelEntry == null)
                {
                    Console.WriteLine($"Entry '{excelEntryName}' not found in the ZIP archive.");
                    return;
                }

                // Open a stream to the entry's data and load it into Aspose.Cells Workbook
                using (Stream entryStream = excelEntry.Open())
                {
                    Workbook workbook = new Workbook(entryStream);

                    // Example operation: read the name of the first worksheet
                    if (workbook.Worksheets.Count > 0)
                    {
                        string firstSheetName = workbook.Worksheets[0].Name;
                        Console.WriteLine($"First worksheet name: {firstSheetName}");
                    }
                    else
                    {
                        Console.WriteLine("The workbook contains no worksheets.");
                    }

                    // (Optional) Save the workbook to a new file to verify it loaded correctly
                    // workbook.Save(@"C:\Data\ExtractedWorkbook.xlsx");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
