// Title: Import JSON data from a web URL into an Aspose.Cells workbook using HttpClient and LoadOptions (C#)
// Common Searches: C# load JSON from a remote URL into Aspose.Cells workbook | How to use Aspose.Cells LoadOptions with a stream to import JSON | Read JSON via HttpClient and convert to Excel with Aspose.Cells | Create an Aspose.Cells workbook from a JSON stream in .NET | Download JSON and save as XLSX using Aspose.Cells in C#
// Tags: Aspose.Cells LoadOptions JSON stream | HttpClient JSON download for workbook | C# import JSON to Excel Aspose.Cells | network JSON to XLSX conversion Aspose | stream-based workbook loading Aspose.Cells

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// This example demonstrates how to download JSON content from a web URL using HttpClient, create LoadOptions with LoadFormat.Json, and load the JSON stream into an Aspose.Cells Workbook. The workbook is then saved as an Excel file, with error handling for network, I/O, and unexpected exceptions.
class Program
{
    static async Task Main(string[] args)
    {
        // URL of the JSON data to load
        const string jsonUrl = "https://example.com/data.json";

        try
        {
            // Initialize HttpClient for network request
            using HttpClient httpClient = new HttpClient();

            // Retrieve the JSON content as a stream
            using Stream jsonStream = await httpClient.GetStreamAsync(jsonUrl);

            // Load the JSON data into an Aspose.Cells Workbook using LoadOptions for JSON format
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Json);
            Workbook workbook = new Workbook(jsonStream, loadOptions);

            // Optional: save the workbook to verify the import
            const string outputPath = "Result.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (HttpRequestException ex)
        {
            Console.Error.WriteLine($"Error retrieving JSON data: {ex.Message}");
        }
        catch (IOException ex)
        {
            Console.Error.WriteLine($"IO error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
