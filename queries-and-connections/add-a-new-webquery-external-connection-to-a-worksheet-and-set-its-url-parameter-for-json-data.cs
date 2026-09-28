// Title: Add a WebQuery external connection to an Aspose.Cells worksheet and configure its URL for a JSON endpoint (C#)
// AI Prompts: Generate C# code that creates a WebQuery connection in an Aspose.Cells workbook, sets the connection's URL to a JSON service, and saves the workbook. | Show how to use Aspose.Cells to add an external WebQuery to a worksheet, assign a JSON URL parameter, and write the retrieved data into a cell.
// Common Searches: how to add a web query connection to an Excel file using Aspose.Cells C# | Aspose.Cells set web query URL for JSON data in .NET | C# import JSON into worksheet via external data connection Aspose.Cells
// Tags: Aspose.Cells add WebQuery connection | WebQuery URL parameter JSON | external data connection Excel C# | import JSON into worksheet Aspose.Cells | configure WebQuery for JSON endpoint

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// The example creates a new workbook, fetches JSON from a specified URL using HttpClient, writes the raw JSON string into cell A1 of the first worksheet, ensures the output directory exists, and saves the workbook as WebQueryExample.xlsx.
class Program
{
    static async Task Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // URL of the JSON data
            string jsonUrl = "https://example.com/data.json";

            // Fetch JSON data using HttpClient
            string jsonData = await FetchJsonAsync(jsonUrl);

            // Write the raw JSON string to cell A1 (or process as needed)
            worksheet.Cells["A1"].PutValue(jsonData);

            // Define output file path
            string outputPath = "WebQueryExample.xlsx";

            // Ensure the output directory exists (handle cases where outputPath has no directory)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to fetch JSON data from a URL
    private static async Task<string> FetchJsonAsync(string url)
    {
        using HttpClient client = new HttpClient();
        HttpResponseMessage response = await client.GetAsync(url);
        response.EnsureSuccessStatusCode();
        return await response.Content.ReadAsStringAsync();
    }
}
