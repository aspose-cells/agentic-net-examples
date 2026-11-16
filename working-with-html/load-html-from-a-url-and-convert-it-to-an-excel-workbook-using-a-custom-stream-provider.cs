// Title: Load HTML from a web URL and convert it to an XLSX workbook using a custom IStreamProvider with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load HTML from a URL with custom stream provider | convert remote HTML page to Excel workbook using HtmlLoadOptions in .NET | download images referenced in HTML during Excel conversion with Aspose.Cells | implement IStreamProvider for external resources in Aspose.Cells HTML import | save HTML stream as .xlsx using Aspose.Cells in C#
// Tags: custom IStreamProvider implementation Aspose.Cells | HtmlLoadOptions with StreamProvider example | load HTML from URL into Workbook | HTML to XLSX conversion handling external images | save workbook as Xlsx using Aspose.Cells

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// The example shows how to implement a custom IStreamProvider that synchronously downloads external resources (such as images) referenced in HTML, fetches an HTML page from a URL with HttpClient, loads the HTML stream into an Aspose.Cells Workbook via HtmlLoadOptions, and saves the result as an XLSX file.
class CustomStreamProvider : IStreamProvider
{
    // Called before any GetStream calls. No initialization needed for this provider.
    public void InitStream(StreamProviderOptions options)
    {
        // No initialization required.
    }

    // Retrieves external resources (e.g., images) referenced in the HTML.
    public Stream GetStream(string uri)
    {
        try
        {
            // Synchronously download the resource and return it as a memory stream.
            using var client = new HttpClient();
            byte[] data = client.GetByteArrayAsync(uri).Result;
            return new MemoryStream(data);
        }
        catch (Exception ex)
        {
            // Return an empty stream on failure to avoid breaking the conversion.
            Console.Error.WriteLine($"Failed to retrieve resource '{uri}': {ex.Message}");
            return new MemoryStream();
        }
    }

    // Called after processing is complete. No cleanup needed for this provider.
    public void CloseStream(StreamProviderOptions options)
    {
        // No cleanup required.
    }
}

class Program
{
    static async Task Main(string[] args)
    {
        try
        {
            // URL of the HTML page to convert.
            const string htmlUrl = "https://example.com/page.html";

            // Download the HTML content as a stream.
            using var httpClient = new HttpClient();
            using Stream htmlStream = await httpClient.GetStreamAsync(htmlUrl);

            // Configure load options to use the custom stream provider.
            var loadOptions = new HtmlLoadOptions
            {
                StreamProvider = new CustomStreamProvider()
            };

            // Load the HTML stream into an Aspose.Cells workbook.
            var workbook = new Workbook(htmlStream, loadOptions);

            // Save the workbook as an Excel file.
            const string outputPath = "ConvertedFromHtml.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
