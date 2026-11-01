// Title: Convert HTML from a Web URL to an Excel workbook using Aspose.Cells and a MemoryStream in C#
// AI Prompts: Write C# code that downloads an HTML page from a specified URL, loads it into an Aspose.Cells Workbook via a MemoryStream, and saves the workbook as an XLSX file. | Refactor the example to use async/await with HttpClient, accept the source URL and destination file path as parameters, and ensure all disposable objects are properly released. | Demonstrate how to configure LoadOptions for HTML format when creating a Workbook from a stream with Aspose.Cells.
// Common Searches: how to load html from a url into Aspose.Cells workbook using memory stream c# | aspocells convert web page to excel file programmatically | c# download html content and save as xlsx with Aspose.Cells | using LoadOptions Html with Aspose.Cells to import html stream
// Tags: Aspose.Cells HTML stream import to workbook | MemoryStream based HTML to XLSX conversion | C# HttpClient fetch HTML for Aspose.Cells processing | LoadOptions Html usage in Aspose.Cells | Save Aspose.Cells workbook in XLSX format

using System;
using System.IO;
using System.Net.Http;
using Aspose.Cells;

// The sample program uses HttpClient to download HTML from a given URL, copies the response into a MemoryStream, loads the HTML into an Aspose.Cells Workbook with LoadOptions set to Html, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        const string htmlUrl = "https://example.com/sample.html";
        const string outputPath = "ConvertedFromHtml.xlsx";

        try
        {
            // Download HTML content from the URL
            using (HttpClient client = new HttpClient())
            {
                HttpResponseMessage response = client.GetAsync(htmlUrl).Result;
                response.EnsureSuccessStatusCode();

                using (MemoryStream htmlStream = new MemoryStream())
                {
                    // Copy response content to a memory stream
                    response.Content.CopyToAsync(htmlStream).Wait();
                    htmlStream.Position = 0; // Reset stream position for reading

                    // Load the HTML into a workbook
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                    using (Workbook workbook = new Workbook(htmlStream, loadOptions))
                    {
                        // Save the workbook as an Excel file
                        workbook.Save(outputPath, SaveFormat.Xlsx);
                        Console.WriteLine($"Workbook saved to '{outputPath}'.");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
