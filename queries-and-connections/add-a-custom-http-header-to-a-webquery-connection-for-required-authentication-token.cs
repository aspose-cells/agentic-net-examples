// Title: Add a custom Authorization header to a WebQuery connection in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# example for adding Authorization header to a WebQuery | How to send an authenticated HTTP GET request and import data into an Excel workbook using Aspose.Cells | C# code to fetch API data with bearer token and write response to a worksheet cell in Aspose.Cells | Using HttpClient with custom headers in Aspose.Cells WebQuery scenario | Save workbook after performing an authenticated web request with Aspose.Cells .NET
// Tags: add Authorization header Aspose.Cells WebQuery | authenticated HttpClient request Aspose.Cells C# | write API response to Excel cell Aspose.Cells | custom header integration Aspose.Cells .NET | save workbook after web query Aspose.Cells

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// Demonstrates how to configure an HttpClient with a custom Authorization (Bearer) header, retrieve data from a protected web API, place the raw response into cell A1 of the first worksheet, and save the workbook as WebQueryWithAuth.xlsx using Aspose.Cells for .NET.
class Program
{
    static async Task Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the URL of the web request
            string requestUrl = "https://api.example.com/data";

            // Prepare HttpClient with custom Authorization header
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Add("Authorization", "Bearer your-auth-token");

                // Send GET request and read response as string
                HttpResponseMessage response = await client.GetAsync(requestUrl);
                response.EnsureSuccessStatusCode();
                string responseData = await response.Content.ReadAsStringAsync();

                // Write the raw response data to cell A1 (or process as needed)
                sheet.Cells["A1"].PutValue(responseData);
            }

            // Define output path
            string outputPath = "WebQueryWithAuth.xlsx";

            // Ensure the output directory exists
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
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
