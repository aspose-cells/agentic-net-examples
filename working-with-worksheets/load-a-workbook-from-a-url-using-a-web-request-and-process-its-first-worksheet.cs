// Title: Load an Excel workbook from a remote URL using HttpClient, modify the first worksheet with Aspose.Cells, and save the updated file
// AI Prompts: Generate C# code that uses HttpClient to download an .xlsx file from a specified URL, streams the response into an Aspose.Cells Workbook, reads the value of cell A1 on the first worksheet, writes a new value to cell B2, and saves the workbook as a new file. | Create a resilient C# routine that, if the HTTP download fails, loads a local Excel file or creates a new workbook, then performs the same first‑worksheet cell updates using Aspose.Cells.
// Common Searches: how to stream an Excel file from a web URL into Aspose.Cells in C# | Aspose.Cells load workbook from HttpClient response stream | C# fallback to local Excel file when remote download fails using Aspose.Cells | update first worksheet cell values after downloading .xlsx with Aspose.Cells | save modified workbook as a new .xlsx file using Aspose.Cells .NET
// Tags: httpclient download excel aspocells | load workbook from stream aspocells .net | fallback to local excel file aspocells | edit first worksheet cell aspocells | save modified workbook aspocells

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// The program attempts to download an .xlsx file from a given URL using HttpClient, loads it into an Aspose.Cells Workbook (or falls back to a local file or a new workbook), reads cell A1 from the first worksheet, writes "Processed" to cell B2, and saves the updated workbook as "Processed.xlsx".
class Program
{
    static async Task Main(string[] args)
    {
        Workbook workbook = null;
        string url = "https://example.com/sample.xlsx";
        string fallbackPath = "sample.xlsx";

        try
        {
            using (HttpClient client = new HttpClient())
            using (HttpResponseMessage response = await client.GetAsync(url))
            {
                if (response.IsSuccessStatusCode)
                {
                    using (Stream stream = await response.Content.ReadAsStreamAsync())
                    {
                        // Load workbook from downloaded stream
                        workbook = new Workbook(stream);
                    }
                }
                else
                {
                    Console.WriteLine($"Warning: Unable to download file (Status {response.StatusCode}). Attempting to load local fallback.");
                }
            }
        }
        catch (HttpRequestException ex)
        {
            Console.WriteLine($"Error downloading file: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }

        // If download failed, try loading a local file if it exists
        if (workbook == null)
        {
            if (File.Exists(fallbackPath))
            {
                try
                {
                    workbook = new Workbook(fallbackPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load local workbook: {ex.Message}");
                    return;
                }
            }
            else
            {
                // Create a new workbook as a last resort
                workbook = new Workbook();
                Console.WriteLine("Created a new workbook because no source file was available.");
            }
        }

        try
        {
            // Access the first worksheet
            Worksheet firstSheet = workbook.Worksheets[0];

            // Read value from cell A1
            string a1Value = firstSheet.Cells["A1"].StringValue;
            Console.WriteLine($"Value of A1: {a1Value}");

            // Write a value to cell B2
            firstSheet.Cells["B2"].PutValue("Processed");

            // Save the modified workbook
            string outputPath = "Processed.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing workbook: {ex.Message}");
        }
    }
}
