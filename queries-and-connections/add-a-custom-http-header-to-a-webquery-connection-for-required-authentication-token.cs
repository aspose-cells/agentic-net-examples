// Title: Add a Bearer authentication header to a WebQuery request and import the response into an Aspose.Cells workbook using C#
// AI Prompts: Generate C# code that sets an Authorization: Bearer header on HttpClient, fetches JSON from a REST endpoint, and writes the result into cell A1 of a new Aspose.Cells workbook. | Show how to create a helper method that accepts a URL and token, performs an authenticated GET request, and returns the response string for use with Aspose.Cells. | Provide a complete example that saves the workbook after loading data from an API that requires a Bearer token, including error handling.
// Common Searches: how to include a bearer token in Aspose.Cells WebQuery C# example | Aspose.Cells fetch data from secured API and write to Excel | C# HttpClient Authorization header with Aspose.Cells workbook import | sample code for authenticated web query using Aspose.Cells | retrieve protected REST data and populate Excel using Aspose.Cells C#
// Tags: bearer token authentication Aspose.Cells WebQuery | C# HttpClient Authorization header Excel import | fetch API data into Aspose.Cells worksheet | save workbook after authenticated web request | custom HTTP header Aspose.Cells connection

using System;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// The example creates a new Workbook, configures HttpClient with an optional Bearer token in the Authorization header, retrieves data from a protected web service, writes the response string to cell A1, and saves the workbook as WebQueryWithHeader.xlsx.
class Program
{
    static async Task Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // URL of the web service providing data
            string url = "https://api.example.com/data";

            // Optional: add authentication header (replace with actual token)
            string authToken = "YOUR_AUTH_TOKEN";

            // Retrieve data from the web service
            string responseData = await GetWebDataAsync(url, authToken);

            // Write the retrieved data to cell A1
            sheet.Cells["A1"].PutValue(responseData);

            // Save the workbook to a file
            string outputPath = "WebQueryWithHeader.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to perform HTTP GET with optional Authorization header
    private static async Task<string> GetWebDataAsync(string requestUrl, string bearerToken = null)
    {
        using (HttpClient client = new HttpClient())
        {
            if (!string.IsNullOrEmpty(bearerToken))
            {
                client.DefaultRequestHeaders.Authorization =
                    new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", bearerToken);
            }

            HttpResponseMessage response = await client.GetAsync(requestUrl);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }
    }
}
