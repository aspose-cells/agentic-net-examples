// Title: Download an Excel workbook from a cloud URL and convert it to CSV with Aspose.Cells in C#
// AI Prompts: Write C# code that fetches an .xlsx file from a specified HTTP URL, saves it to a temporary path, and uses Aspose.Cells ConversionUtility to export the workbook as a CSV file. | Enhance the program to add retry logic with exponential back‑off for the download step and log success or failure of the Excel‑to‑CSV conversion using Aspose.Cells.
// Common Searches: c# Aspose.Cells convert xlsx from remote URL to csv | how to download an Excel file from cloud storage and export to csv using Aspose.Cells | Aspose.Cells ConversionUtility example for converting a downloaded workbook to csv in .NET | download spreadsheet with HttpClient and convert to csv using Aspose.Cells | temporary file cleanup after Aspose.Cells conversion c#
// Tags: Aspose.Cells ConversionUtility xlsx to csv | download workbook with HttpClient C# | temporary file handling for Aspose.Cells conversion | error handling for HTTP download Aspose.Cells | cloud storage Excel to CSV conversion .NET

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells.Utility;

// The example downloads an .xlsx workbook from a cloud URL using HttpClient, stores it in a temporary file, converts it to CSV with Aspose.Cells ConversionUtility, outputs the CSV location, and then removes the temporary workbook.
class Program
{
    static async Task Main()
    {
        // URL of the workbook stored in cloud storage
        string cloudUrl = "https://example.com/sample.xlsx";

        // Temporary local path to store the downloaded workbook
        string tempWorkbookPath = Path.Combine(Path.GetTempPath(), "tempWorkbook.xlsx");

        // Desired output CSV file path
        string csvOutputPath = Path.Combine(Environment.CurrentDirectory, "output.csv");

        try
        {
            // Download the workbook from cloud storage
            using (HttpClient httpClient = new HttpClient())
            {
                HttpResponseMessage response;
                try
                {
                    response = await httpClient.GetAsync(cloudUrl);
                }
                catch (HttpRequestException ex)
                {
                    Console.WriteLine($"Error downloading workbook: {ex.Message}");
                    return;
                }

                if (!response.IsSuccessStatusCode)
                {
                    Console.WriteLine($"Failed to download workbook. Status code: {(int)response.StatusCode} {response.ReasonPhrase}");
                    return;
                }

                await using (FileStream fileStream = new FileStream(tempWorkbookPath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    await response.Content.CopyToAsync(fileStream);
                }
            }

            // Ensure the downloaded file exists before conversion
            if (!File.Exists(tempWorkbookPath))
            {
                Console.WriteLine("Downloaded workbook file not found.");
                return;
            }

            // Convert the downloaded workbook to CSV using Aspose.Cells ConversionUtility
            try
            {
                ConversionUtility.Convert(tempWorkbookPath, csvOutputPath);
                Console.WriteLine($"Workbook successfully converted to CSV at: {csvOutputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
        finally
        {
            // Clean up the temporary workbook file
            if (File.Exists(tempWorkbookPath))
            {
                try
                {
                    File.Delete(tempWorkbookPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unable to delete temporary file: {ex.Message}");
                }
            }
        }
    }
}
