using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsJsonFromNetwork
{
    class Program
    {
        static async Task Main(string[] args)
        {
            // URL of the JSON resource on the network (replace with a valid endpoint)
            string jsonUrl = "https://jsonplaceholder.typicode.com/todos/1";

            string jsonContent;

            using (HttpClient httpClient = new HttpClient())
            {
                try
                {
                    jsonContent = await httpClient.GetStringAsync(jsonUrl);
                }
                catch (HttpRequestException)
                {
                    // Fallback JSON in case the request fails
                    jsonContent = @"{ ""userId"": 1, ""id"": 1, ""title"": ""delectus aut autem"", ""completed"": false }";
                }
            }

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure layout options (optional)
            JsonLayoutOptions layoutOptions = new JsonLayoutOptions
            {
                ArrayAsTable = true
            };

            // Import the JSON string into the worksheet starting at cell A1 (row 0, column 0)
            JsonUtility.ImportData(jsonContent, worksheet.Cells, 0, 0, layoutOptions);

            // Save the resulting workbook to an Excel file
            string outputPath = "NetworkJsonOutput.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"JSON data has been imported and saved to '{outputPath}'.");
        }
    }
}