// Title: Download an Excel workbook from a cloud URL and save it as CSV using Aspose.Cells for .NET
// Common Searches: C# code to download an .xlsx file from a web URL and convert it to CSV with Aspose.Cells | How to load an Excel workbook from HttpClient stream in Aspose.Cells .NET | Aspose.Cells example for converting a remote workbook to CSV in C# | Fallback to local Excel file when cloud download fails using Aspose.Cells | Save first worksheet as CSV using Aspose.Cells SaveFormat in .NET | Convert Excel to CSV without saving intermediate file using memory stream Aspose.Cells
// Tags: download workbook via HttpClient Aspose.Cells | load workbook from memory stream C# | fallback local Excel loading Aspose.Cells | save worksheet as CSV SaveFormat.Csv | convert remote .xlsx to .csv Aspose.Cells | Aspose.Cells cloud to CSV conversion

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// The sample demonstrates how to download an Excel workbook from a cloud URL using HttpClient, optionally fall back to a local file or create a new workbook, and then export the first worksheet to CSV with Aspose.Cells' Save method and SaveFormat.Csv.
class Program
{
    static async Task Main()
    {
        // URL of the workbook stored in the cloud
        string cloudFileUrl = "https://example.com/path/to/workbook.xlsx";

        // Path to a local fallback workbook (optional)
        string localFallbackPath = "fallback.xlsx";

        Workbook workbook = null;

        try
        {
            // Attempt to download the workbook from the cloud
            using (HttpClient httpClient = new HttpClient())
            using (HttpResponseMessage response = await httpClient.GetAsync(cloudFileUrl))
            {
                if (response.IsSuccessStatusCode)
                {
                    using (Stream downloadStream = await response.Content.ReadAsStreamAsync())
                    using (MemoryStream memoryStream = new MemoryStream())
                    {
                        await downloadStream.CopyToAsync(memoryStream);
                        memoryStream.Position = 0; // Reset stream position for reading
                        workbook = new Workbook(memoryStream);
                    }
                }
                else
                {
                    Console.WriteLine($"Failed to download workbook. HTTP {(int)response.StatusCode} {response.ReasonPhrase}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Exception during download: {ex.Message}");
        }

        // If download failed, try loading a local fallback workbook
        if (workbook == null)
        {
            if (File.Exists(localFallbackPath))
            {
                try
                {
                    workbook = new Workbook(localFallbackPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load local fallback workbook: {ex.Message}");
                }
            }
            else
            {
                // No source workbook available; create a new empty workbook
                workbook = new Workbook();
                // Optionally add sample data to the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample");
                sheet.Cells["B1"].PutValue("Data");
            }
        }

        // Convert and save the first worksheet as CSV
        try
        {
            string outputCsvPath = "output.csv";
            workbook.Save(outputCsvPath, SaveFormat.Csv);
            Console.WriteLine($"Workbook saved as CSV to '{outputCsvPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving CSV: {ex.Message}");
        }
    }
}
