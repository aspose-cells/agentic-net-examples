// Title: Load an Excel workbook from a ZIP file by streaming the entry into Aspose.Cells Workbook constructor (C#)
// Common Searches: C# Aspose.Cells load Excel file directly from zip entry stream | How to read a workbook inside a .zip without extracting to disk using Aspose.Cells | Open Excel workbook from compressed archive with Aspose.Cells in .NET | Stream an Excel file from a ZipArchive to Aspose.Cells Workbook constructor | Aspose.Cells example for loading workbook from zip archive in C#
// Tags: Aspose.Cells load workbook from zip stream | C# ZipArchive extract Excel entry to Workbook | Workbook constructor stream input Aspose.Cells | read Excel inside zip without temporary file | Aspose.Cells open workbook from compressed archive

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

// The sample demonstrates how to verify a ZIP archive, locate an Excel entry, obtain its stream, and instantiate an Aspose.Cells Workbook directly from that stream in C#. The workbook is then saved to a new file, with comprehensive error handling for missing files or entries.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the ZIP archive containing the Excel file
            string zipPath = "sample.zip";

            // Verify that the ZIP file exists before attempting to open it
            if (!File.Exists(zipPath))
            {
                Console.WriteLine($"Error: ZIP archive '{zipPath}' not found.");
                return;
            }

            // Name of the Excel file inside the ZIP archive
            string entryName = "Workbook.xlsx";

            // Load the workbook from the ZIP archive
            Workbook workbook = LoadWorkbookFromZip(zipPath, entryName);
            if (workbook == null)
            {
                Console.WriteLine("Failed to load workbook from ZIP archive.");
                return;
            }

            // Example usage: save the extracted workbook to a new file
            string outputPath = "ExtractedWorkbook.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook extracted and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    static Workbook LoadWorkbookFromZip(string zipFilePath, string entryFileName)
    {
        // Open the ZIP archive for reading
        using (FileStream zipStream = new FileStream(zipFilePath, FileMode.Open, FileAccess.Read))
        using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Read))
        {
            // Locate the specific entry (Excel file) inside the archive
            ZipArchiveEntry entry = archive.GetEntry(entryFileName);
            if (entry == null)
            {
                Console.WriteLine($"Error: Entry '{entryFileName}' not found in the ZIP archive.");
                return null;
            }

            // Open the entry's stream and pass it to Aspose.Cells Workbook constructor
            using (Stream entryStream = entry.Open())
            {
                // The Workbook constructor reads the stream and creates the workbook object
                return new Workbook(entryStream);
            }
        }
    }
}
