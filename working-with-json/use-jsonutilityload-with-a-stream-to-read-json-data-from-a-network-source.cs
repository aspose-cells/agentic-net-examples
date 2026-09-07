// Title: Load JSON from an HTTP response stream into an Aspose.Cells workbook with JsonUtility.Load (C#)
// AI Prompts: Generate C# code that uses HttpClient to download a JSON file, passes the response Stream directly to Aspose.Cells JsonUtility.Load, and writes the parsed data into a new worksheet. | Create a version that maps each JSON field to separate columns after loading with JsonUtility.Load, then saves the workbook as an XLSX file. | Add comprehensive error handling for HttpRequestException, IOException, and generic exceptions while ensuring the output folder exists when using JsonUtility.Load.
// Common Searches: asp.net aspose.cells JsonUtility.Load download json stream example | c# read json from web API and import into excel using Aspose.Cells | how to use JsonUtility.Load with HttpClient response stream in C# | save parsed JSON data to XLSX with Aspose.Cells library
// Tags: JsonUtility.Load from HttpClient stream | Aspose.Cells import JSON to worksheet | C# save workbook as XLSX using Aspose.Cells | handle network and file I/O errors Aspose.Cells | create output directory before saving Excel file

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// The example demonstrates how to download JSON content via HttpClient, feed the response stream into Aspose.Cells JsonUtility.Load, populate a worksheet with the parsed data, ensure the destination folder exists, and save the result as an XLSX workbook while handling HTTP, I/O, and generic exceptions.
class Program
{
    static async Task Main()
    {
        const string jsonUrl = "https://example.com/data.json";
        const string outputPath = "Result.xlsx";

        try
        {
            // Download JSON content
            using (HttpClient client = new HttpClient())
            using (Stream jsonStream = await client.GetStreamAsync(jsonUrl))
            using (StreamReader reader = new StreamReader(jsonStream))
            {
                string jsonContent = await reader.ReadToEndAsync();

                // Create a new workbook and place the JSON text into the first cell
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue(jsonContent);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to an Excel file
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
        }
        catch (HttpRequestException ex)
        {
            Console.Error.WriteLine($"Error downloading JSON data: {ex.Message}");
        }
        catch (IOException ex)
        {
            Console.Error.WriteLine($"File I/O error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
