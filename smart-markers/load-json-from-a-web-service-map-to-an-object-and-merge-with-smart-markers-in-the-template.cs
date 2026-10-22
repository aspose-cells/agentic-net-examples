// Title: Retrieve JSON from a REST API, deserialize to C# objects, and merge into an Excel template using Aspose.Cells smart markers
// AI Prompts: Call a web service to obtain JSON, deserialize it into a List<Person>, assign the list to the "People" data source of WorkbookDesigner, and process the smart markers to produce the merged workbook. | Write a C# program that loads an Excel file containing smart markers, binds data fetched from an HTTP endpoint, and saves the populated workbook using Aspose.Cells.
// Common Searches: asp.net core fetch json and fill Aspose.Cells smart markers template | c# Aspose.Cells WorkbookDesigner populate smart markers from REST API response | how to bind a list of objects to smart markers in an Excel template using Aspose.Cells | merge json data into excel using Aspose.Cells smart markers c# example | load excel template with smart markers and replace with data from web service
// Tags: Aspose.Cells WorkbookDesigner JSON data source | C# smart markers Excel template merge | deserialize REST API response to list for Excel | populate Excel smart markers from web service | process smart markers with dynamic data in .NET

using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsSmartMarkerExample
{
    // Data model matching the JSON structure.
    // The program fetches JSON from a specified REST endpoint, deserializes it into a List<Person>, loads an Excel template that contains Aspose.Cells smart markers, sets the "People" data source for WorkbookDesigner, processes the markers to merge the data, and saves the resulting workbook while handling HTTP, JSON, and file errors.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        // Add other properties as needed.
    }

    class Program
    {
        // Entry point – async to allow awaiting the web request.
        static async Task Main(string[] args)
        {
            // URL of the web service returning JSON.
            const string jsonUrl = "https://example.com/api/people";

            // Path to the Excel template that contains smart markers.
            const string templatePath = @"C:\Templates\SmartMarkerTemplate.xlsx";

            // Path where the merged workbook will be saved.
            const string outputPath = @"C:\Output\MergedResult.xlsx";

            try
            {
                // 1. Retrieve JSON from the web service.
                string jsonContent;
                using (HttpClient httpClient = new HttpClient())
                {
                    jsonContent = await httpClient.GetStringAsync(jsonUrl);
                }

                // 2. Deserialize JSON into a list of Person objects.
                List<Person> people = JsonSerializer.Deserialize<List<Person>>(jsonContent, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                }) ?? new List<Person>();

                // 3. Verify that the template file exists before loading.
                if (!File.Exists(templatePath))
                {
                    throw new FileNotFoundException($"Template file not found: {templatePath}");
                }

                // Load the Excel template.
                Workbook workbook = new Workbook(templatePath);

                // 4. Use WorkbookDesigner to process smart markers.
                WorkbookDesigner designer = new WorkbookDesigner(workbook);
                // The data source name "People" must match the name used in the template.
                designer.SetDataSource("People", people);
                designer.Process();

                // Ensure the output directory exists.
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // 5. Save the merged workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook merged and saved to: {outputPath}");
            }
            catch (HttpRequestException httpEx)
            {
                Console.Error.WriteLine($"Error retrieving JSON data: {httpEx.Message}");
            }
            catch (JsonException jsonEx)
            {
                Console.Error.WriteLine($"Error deserializing JSON: {jsonEx.Message}");
            }
            catch (FileNotFoundException fileEx)
            {
                Console.Error.WriteLine(fileEx.Message);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
