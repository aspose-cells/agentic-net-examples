// Title: Convert HTML from a Web URL to an XLSX Workbook Using Aspose.Cells and a MemoryStream in C#
// AI Prompts: Write C# code that downloads an HTML page with HttpClient, loads the content into a MemoryStream, creates an Aspose.Cells Workbook via HtmlLoadOptions, and saves it as an .xlsx file. | Modify the example to accept the source URL and output file path as command‑line arguments, while still using a stream‑based conversion with Aspose.Cells. | Add comprehensive error handling for network failures, invalid HTML, and ensure all disposable objects are wrapped in using statements during the HTML‑to‑Excel conversion.
// Common Searches: aspnet convert remote html page to excel workbook using aspose.cells | c# load html from url into workbook without saving temporary file | how to use HtmlLoadOptions with a memory stream in Aspose.Cells | download html with HttpClient and export to xlsx in .net core | stream based html to xlsx conversion using Aspose.Cells C# example
// Tags: html to xlsx conversion using Aspose.Cells MemoryStream | load HTML stream into Aspose.Cells Workbook | Aspose.Cells HtmlLoadOptions with HttpClient | download web page as MemoryStream for Excel export | command line parameters for HTML to Excel conversion C#

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// The sample downloads HTML from a specified URL using HttpClient, copies it into a MemoryStream, loads it into an Aspose.Cells Workbook via HtmlLoadOptions, and saves the workbook as an XLSX file.
class Program
{
    static async Task Main(string[] args)
    {
        try
        {
            // URL of the HTML page to convert
            const string url = "https://example.com/sample.html";

            // Download the HTML content into a memory stream
            using (HttpClient httpClient = new HttpClient())
            using (Stream remoteStream = await httpClient.GetStreamAsync(url))
            using (MemoryStream htmlMemory = new MemoryStream())
            {
                await remoteStream.CopyToAsync(htmlMemory);
                htmlMemory.Position = 0; // reset to beginning for reading

                // Load the HTML into an Aspose.Cells workbook directly from the stream
                var loadOptions = new HtmlLoadOptions(); // default loading options
                Workbook workbook = new Workbook(htmlMemory, loadOptions);

                // Save the workbook as an Excel file
                const string outputPath = "ConvertedFromHtml.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
