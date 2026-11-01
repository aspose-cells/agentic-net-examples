// Title: Download an Excel workbook from a cloud URL and convert the first worksheet to CSV using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells download Excel file from HTTPS URL and save as CSV | How to load an .xlsx from a web address into an Aspose.Cells Workbook in .NET | Convert remote Excel workbook to CSV with Aspose.Cells using HttpClient stream | Aspose.Cells .NET save first worksheet as CSV after downloading from cloud storage | Fallback to local Excel file when HTTP download fails with Aspose.Cells
// Tags: download Excel via HttpClient Aspose.Cells | load workbook from memory stream Aspose.Cells | save worksheet as CSV Aspose.Cells | ensure output directory exists before saving CSV C# | fallback to local file on download error Aspose.Cells

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// The example uses HttpClient to download an .xlsx file from a cloud URL, copies it into a MemoryStream, loads it into an Aspose.Cells Workbook, optionally falls back to a local file if the download fails, creates the target directory if needed, and saves the first worksheet as a CSV file using SaveFormat.Csv, with comprehensive error handling.
class Program
{
    static async Task Main(string[] args)
    {
        // URL of the Excel file stored in cloud storage
        string excelUrl = "https://example.com/path/to/workbook.xlsx";

        // Destination path for the CSV output
        string csvPath = "output.csv";

        Workbook workbook = null;

        try
        {
            // Attempt to download the Excel file from the URL
            using (HttpClient httpClient = new HttpClient())
            using (Stream excelStream = await httpClient.GetStreamAsync(excelUrl))
            using (MemoryStream memoryStream = new MemoryStream())
            {
                // Copy the downloaded stream to a memory stream (required for Aspose.Cells)
                await excelStream.CopyToAsync(memoryStream);
                memoryStream.Position = 0; // Reset position to the beginning

                // Load the workbook from the memory stream
                workbook = new Workbook(memoryStream);
            }
        }
        catch (HttpRequestException ex)
        {
            Console.WriteLine($"Failed to download workbook from URL: {ex.Message}");
            // Fallback: try to load a local file with the same name if it exists
            string localPath = "workbook.xlsx";
            if (File.Exists(localPath))
            {
                try
                {
                    workbook = new Workbook(localPath);
                }
                catch (Exception loadEx)
                {
                    Console.WriteLine($"Failed to load local workbook: {loadEx.Message}");
                }
            }
            else
            {
                Console.WriteLine("No local fallback workbook found.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error while downloading workbook: {ex.Message}");
        }

        if (workbook == null)
        {
            Console.WriteLine("Workbook could not be loaded. Exiting.");
            return;
        }

        try
        {
            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(csvPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the first worksheet as CSV
            workbook.Worksheets[0].PageSetup.PrintArea = ""; // Ensure entire sheet is saved
            workbook.Save(csvPath, SaveFormat.Csv);

            Console.WriteLine($"Workbook has been converted to CSV and saved at: {csvPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during CSV conversion or saving: {ex.Message}");
        }
    }
}
