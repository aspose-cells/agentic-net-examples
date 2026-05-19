using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

class Program
{
    static async Task Main()
    {
        // URL of the source Excel file
        const string sourceUrl = "https://example.com/sample.xlsx";

        // Path where the compressed workbook will be saved
        const string destinationPath = "output.xlsx";

        try
        {
            // Ensure the destination directory exists
            string destDir = Path.GetDirectoryName(Path.GetFullPath(destinationPath));
            if (!Directory.Exists(destDir))
                Directory.CreateDirectory(destDir);

            // Download the workbook using HttpClient
            using (HttpClient httpClient = new HttpClient())
            {
                HttpResponseMessage response = await httpClient.GetAsync(sourceUrl);
                response.EnsureSuccessStatusCode(); // Throws if status is not 2xx

                using (Stream networkStream = await response.Content.ReadAsStreamAsync())
                {
                    // Load the workbook from the downloaded stream
                    Workbook workbook = new Workbook(networkStream);

                    // Configure OOXML save options with maximum compression (Level9)
                    OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                    {
                        CompressionType = OoxmlCompressionType.Level9
                    };

                    // Save the workbook to the specified file using the configured options
                    workbook.Save(destinationPath, saveOptions);
                }
            }

            Console.WriteLine($"Workbook saved successfully to '{destinationPath}'.");
        }
        catch (HttpRequestException ex)
        {
            Console.Error.WriteLine($"Error downloading the file: {ex.Message}");
        }
        catch (FileNotFoundException ex)
        {
            Console.Error.WriteLine($"File not found: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}