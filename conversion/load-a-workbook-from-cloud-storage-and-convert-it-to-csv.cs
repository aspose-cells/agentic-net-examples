// Title: C# – Download an Excel file from a cloud URL and convert it to CSV with Aspose.Cells
// Description: This example shows how to asynchronously download an Excel workbook from cloud storage, load it with Aspose.Cells using default LoadOptions, save it as a CSV file, and clean up the temporary file.
// Keywords: Aspose.Cells | C# download Excel | cloud storage Excel file | convert Excel to CSV | LoadOptions auto detect | SaveFormat.Csv | async file download .NET | temporary file cleanup | CSV conversion .NET
// Common Searches: download Excel from URL and convert to CSV C# | Aspose.Cells convert remote workbook to CSV | load Excel file from cloud storage using Aspose.Cells | C# example for Excel to CSV conversion with Aspose | how to save Aspose.Cells workbook as CSV
// Developer Intent: Retrieve an Excel workbook hosted in cloud storage, load it with Aspose.Cells, and export it as a CSV file in a .NET application.
// Use Cases: Automate nightly conversion of cloud‑hosted Excel reports to CSV for downstream analytics pipelines. | Build a web API that accepts a file URL, transforms the workbook to CSV, and returns the result to callers. | Integrate Excel‑to‑CSV conversion into ETL jobs that pull source files from SaaS storage services.
// AI Prompts: Write C# code that downloads an Excel file from a given cloud URL, loads it with Aspose.Cells, and saves it as CSV with proper error handling. | Create a reusable method in .NET that takes a cloud file link and an output path, performs the download, conversion to CSV, and deletes temporary files. | Explain how to configure Aspose.Cells LoadOptions for specific Excel formats when converting a cloud‑based workbook to CSV.

using System;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    // This example shows how to asynchronously download an Excel workbook from cloud storage, load it with Aspose.Cells using default LoadOptions, save it as a CSV file, and clean up the temporary file.
    class Program
    {
        static async Task Main(string[] args)
        {
            // URL of the Excel file stored in cloud storage
            string cloudFileUrl = "https://example.com/files/sample.xlsx";

            // Temporary local path to download the Excel file
            string tempExcelPath = Path.Combine(Path.GetTempPath(), "temp_downloaded.xlsx");

            // Desired output CSV file path
            string outputCsvPath = "converted_output.csv";

            // Download the Excel file from cloud storage
            await DownloadFileAsync(cloudFileUrl, tempExcelPath);

            // Create load options (auto-detect format or specify if known)
            LoadOptions loadOptions = new LoadOptions(); // default auto detection

            // Load the workbook from the downloaded file using the constructor with LoadOptions
            Workbook workbook = new Workbook(tempExcelPath, loadOptions);

            // Save the workbook as CSV using the Save method with SaveFormat.Csv
            workbook.Save(outputCsvPath, SaveFormat.Csv);

            // Clean up the temporary Excel file
            if (File.Exists(tempExcelPath))
            {
                File.Delete(tempExcelPath);
            }

            Console.WriteLine($"Conversion completed. CSV saved to: {outputCsvPath}");
        }

        // Helper method to download a file from a URL to a local path
        private static async Task DownloadFileAsync(string url, string destinationPath)
        {
            using (HttpClient client = new HttpClient())
            using (HttpResponseMessage response = await client.GetAsync(url))
            using (Stream contentStream = await response.Content.ReadAsStreamAsync())
            using (FileStream fileStream = new FileStream(destinationPath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                await contentStream.CopyToAsync(fileStream);
            }
        }
    }
}
