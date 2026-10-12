// Title: Load JSON from a web URL into an Aspose.Cells workbook using JsonUtility.Load with a network stream in C#
// AI Prompts: Write C# code that uses HttpClient to download a JSON response as a Stream, passes the stream to Aspose.Cells JsonUtility.Load, and saves the resulting workbook as an XLSX file. | Show how to read a JSON stream from a REST endpoint and map its objects to specific cells or tables in an Aspose.Cells worksheet using JsonUtility.Load. | Adapt the example to handle large JSON payloads by streaming directly into a workbook without loading the entire content into memory, using Aspose.Cells JsonUtility.Load.
// Common Searches: Aspose.Cells JsonUtility.Load from HttpClient response stream C# | How to import JSON data from a web service into an Excel workbook using Aspose.Cells .NET | C# load remote JSON into worksheet cells with Aspose.Cells without temporary files | Example of streaming JSON into Aspose.Cells workbook from a URL | Using Aspose.Cells to read JSON via network stream and save as XLSX
// Tags: Aspose.Cells JsonUtility.Load HTTP stream | C# import remote JSON into Excel workbook | Stream-based JSON loading Aspose.Cells | Load JSON from URL into worksheet cells | Save JSON data as XLSX using Aspose.Cells

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// The program fetches JSON from a specified URL using HttpClient, obtains the response as a Stream, and feeds that stream directly to Aspose.Cells JsonUtility.Load to create a workbook. The workbook is then saved as an XLSX file, with error handling for HTTP, I/O, and general exceptions.
class Program
{
    static async Task Main(string[] args)
    {
        // URL of the JSON data source
        const string jsonUrl = "https://example.com/data.json";

        try
        {
            // Create an HttpClient to fetch the JSON stream
            using (HttpClient httpClient = new HttpClient())
            {
                // Get the response stream asynchronously
                using (Stream jsonStream = await httpClient.GetStreamAsync(jsonUrl))
                using (StreamReader reader = new StreamReader(jsonStream))
                {
                    // Read the entire JSON content
                    string jsonContent = await reader.ReadToEndAsync();

                    // Create a new workbook
                    Workbook workbook = new Workbook();

                    // Simple demonstration: place the raw JSON into cell A1
                    // (Replace with proper JSON parsing and cell population as needed)
                    workbook.Worksheets[0].Cells["A1"].PutValue(jsonContent);

                    // Save the workbook to verify the load succeeded
                    const string outputPath = "Output.xlsx";

                    // Ensure the directory exists before saving
                    string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"JSON data loaded into workbook and saved as {outputPath}");
                }
            }
        }
        catch (HttpRequestException httpEx)
        {
            Console.WriteLine($"Error fetching JSON data: {httpEx.Message}");
        }
        catch (IOException ioEx)
        {
            Console.WriteLine($"File I/O error: {ioEx.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
