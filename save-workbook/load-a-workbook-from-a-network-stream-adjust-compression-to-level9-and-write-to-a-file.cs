// Title: Download an Excel workbook via HttpClient, set OOXML Level9 compression, and save locally with Aspose.Cells (C#)
// AI Prompts: Write C# code that uses HttpClient to download an .xlsx file, loads it into an Aspose.Cells Workbook, and saves it with OoxmlSaveOptions.CompressionType set to Level9. | Create error‑handling logic that falls back to creating a new empty Workbook when the download fails, then saves the workbook using maximum OOXML compression. | Show how to configure OoxmlSaveOptions for Level9 compression and pass it to Workbook.Save in Aspose.Cells.
// Common Searches: aspnet download excel file with HttpClient and save using Aspose.Cells compression level9 | how to apply maximum OOXML compression when saving a workbook in C# Aspose.Cells | fallback to empty workbook if network stream download fails Aspose.Cells example
// Tags: HttpClient download Excel file Aspose.Cells | Level9 OOXML compression Aspose.Cells | save workbook with maximum compression C# | load workbook from network stream Aspose.Cells | fallback to empty workbook on download error

using System;
using System.IO;
using System.Net.Http;
using Aspose.Cells;

// This example downloads an .xlsx file via HttpClient, loads it into an Aspose.Cells Workbook (or creates a new one on failure), configures OoxmlSaveOptions with Level9 compression, and saves the workbook to a local file.
class LoadCompressSave
{
    static void Main()
    {
        // URL of the Excel file to download
        string fileUrl = "https://example.com/sample.xlsx";

        // Destination file path
        string outputPath = "CompressedOutput.xlsx";

        Workbook workbook = null;

        // Attempt to download the workbook; fall back to a new workbook on failure
        try
        {
            using (HttpClient client = new HttpClient())
            using (Stream networkStream = client.GetStreamAsync(fileUrl).Result)
            {
                workbook = new Workbook(networkStream);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to download workbook: {ex.Message}");
            Console.WriteLine("Creating a new empty workbook instead.");
            workbook = new Workbook(); // creates a default workbook with one worksheet
        }

        // Configure OOXML save options with maximum compression (Level9)
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
        {
            CompressionType = OoxmlCompressionType.Level9
        };

        // Save the workbook to a file using the configured options
        try
        {
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine("Workbook saved to: " + Path.GetFullPath(outputPath));
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
