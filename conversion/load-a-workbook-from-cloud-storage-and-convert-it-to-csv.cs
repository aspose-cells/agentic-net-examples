using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsCloudToCsv
{
    class Program
    {
        static async Task Main(string[] args)
        {
            // URL of the Excel file stored in cloud storage
            string cloudFileUrl = "https://example.com/path/to/your/file.xlsx";

            // Local path for the resulting CSV file
            string csvOutputPath = "converted.csv";

            try
            {
                // Validate URL
                if (string.IsNullOrWhiteSpace(cloudFileUrl))
                    throw new ArgumentException("Cloud file URL is not provided.");

                using HttpClient httpClient = new HttpClient();

                // Download the Excel file
                using HttpResponseMessage response = await httpClient.GetAsync(cloudFileUrl);
                response.EnsureSuccessStatusCode();

                await using Stream excelStream = await response.Content.ReadAsStreamAsync();

                // Load options for XLSX format
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

                // Load workbook from stream
                Workbook workbook = new Workbook(excelStream, loadOptions);

                // Save as CSV
                workbook.Save(csvOutputPath, SaveFormat.Csv);

                Console.WriteLine($"Workbook downloaded from cloud and saved as CSV to '{csvOutputPath}'.");
            }
            catch (HttpRequestException ex)
            {
                Console.Error.WriteLine($"Error downloading file: {ex.Message}");
            }
            catch (FileNotFoundException ex)
            {
                Console.Error.WriteLine($"File not found: {ex.Message}");
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
}