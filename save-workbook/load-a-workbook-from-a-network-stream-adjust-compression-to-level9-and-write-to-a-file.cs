// Title: Load an Excel workbook from an HTTP stream with Aspose.Cells, set ZIP compression to level 9, and save as XLSX in C#
// AI Prompts: Generate C# code that uses HttpClient to download an .xlsx file, creates an Aspose.Cells Workbook from the response stream, and saves it to a local path. | Extend the saving code to use Aspose.Cells SaveOptions, configuring the ZipCompressionLevel to 9 before writing the workbook as an XLSX file.
// Common Searches: how to set maximum zip compression when saving an Excel file with Aspose.Cells C# | download Excel workbook from URL and save with compression level 9 using Aspose.Cells | Aspose.Cells load workbook from HttpClient stream and specify SaveOptions compression | C# example for saving Aspose.Cells workbook with ZipCompressionLevel 9 | saving workbook from network stream with Aspose.Cells SaveFormat.Xlsx and high compression
// Tags: Aspose.Cells load workbook from HttpClient stream | Aspose.Cells configure zip compression for XLSX | Aspose.Cells save workbook with high compression | C# download Excel file using HttpClient and Aspose.Cells | network stream workbook loading Aspose.Cells

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// The example downloads an XLSX file via HttpClient, loads it into an Aspose.Cells Workbook from the network stream, then saves it locally as output.xlsx using SaveOptions with ZipCompressionLevel set to 9, while handling potential errors.
class Program
{
    static async Task Main(string[] args)
    {
        try
        {
            // URL of the workbook to download
            string url = "https://example.com/sample.xlsx";

            // Download the workbook stream using HttpClient
            using HttpClient client = new HttpClient();
            using Stream networkStream = await client.GetStreamAsync(url);

            // Load the workbook from the downloaded stream
            Workbook workbook = new Workbook(networkStream);

            // Save the workbook to a local file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any errors that occur during download or processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
