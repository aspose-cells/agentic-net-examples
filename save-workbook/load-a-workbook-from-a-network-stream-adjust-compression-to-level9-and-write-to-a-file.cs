// Title: Download an Excel file via HttpClient, load it from a network stream, and save as XLSB with Level9 compression using Aspose.Cells for .NET
// Common Searches: c# download xlsx from url and save as xlsb with level9 compression using Aspose.Cells | how to set OoxmlCompressionType.Level9 when saving a workbook to XLSB in Aspose.Cells | load workbook from HttpClient stream Aspose.Cells example | save workbook as compressed XLSB file in .NET with Aspose.Cells | Aspose.Cells XlsbSaveOptions compression level9 sample code
// Tags: Aspose.Cells HttpClient stream workbook loading | XlsbSaveOptions OoxmlCompressionType Level9 configuration | save workbook to XLSB with maximum compression Aspose.Cells | network stream Excel file handling Aspose.Cells | compressed XLSB output using Aspose.Cells .NET

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

// The example downloads an Excel workbook from a URL using HttpClient, loads it directly from the network stream into an Aspose.Cells Workbook, configures XlsbSaveOptions with OoxmlCompressionType.Level9 for maximum compression, and saves the workbook as an XLSB file to a specified path.
class Program
{
    static async Task Main(string[] args)
    {
        // URL of the workbook to download. Can be passed as first argument.
        string workbookUrl = args.Length > 0 ? args[0] : "https://example.com/sample.xlsx";

        // Destination file path for the saved XLSB file. Can be passed as second argument.
        string outputPath = args.Length > 1 ? args[1] : "output.xlsb";

        try
        {
            // Download the workbook into a network stream.
            using HttpClient httpClient = new HttpClient();
            using Stream networkStream = await httpClient.GetStreamAsync(workbookUrl);

            // Load the workbook from the stream.
            using Workbook workbook = new Workbook(networkStream);

            // Configure XLSB save options with maximum compression (Level9).
            XlsbSaveOptions saveOptions = new XlsbSaveOptions
            {
                CompressionType = OoxmlCompressionType.Level9
            };

            // Save the workbook to the specified file using the configured options.
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook successfully saved to '{outputPath}' with compression Level9.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
