using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the ZIP archive and the Excel file inside it
        string zipPath = "sample.zip";
        string excelEntryName = "sample.xlsx";

        // Verify that the ZIP file exists before attempting to open it
        if (!File.Exists(zipPath))
        {
            Console.WriteLine($"ZIP file not found: {zipPath}");
            return;
        }

        try
        {
            // Open the ZIP archive for reading
            using (FileStream zipStream = new FileStream(zipPath, FileMode.Open, FileAccess.Read))
            using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Read))
            {
                // Locate the Excel entry
                ZipArchiveEntry entry = archive.GetEntry(excelEntryName);
                if (entry == null)
                {
                    Console.WriteLine($"Entry '{excelEntryName}' not found in the ZIP archive.");
                    return;
                }

                // Open the entry as a stream and load the workbook
                using (Stream entryStream = entry.Open())
                {
                    Workbook workbook = new Workbook(entryStream);

                    // Example: display the name of the first worksheet
                    Worksheet firstSheet = workbook.Worksheets[0];
                    Console.WriteLine("First worksheet name: " + firstSheet.Name);

                    // Save the extracted workbook to a file
                    string outputPath = "ExtractedWorkbook.xlsx";
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved to: {outputPath}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}