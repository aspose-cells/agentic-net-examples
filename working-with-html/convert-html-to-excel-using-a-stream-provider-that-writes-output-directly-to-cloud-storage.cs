// Title: Convert HTML to Excel in C# and upload via stream to Azure, AWS or GCP using Aspose.Cells
// Description: Load an HTML file into an Aspose.Cells Workbook, save it as XLSX to a MemoryStream, and upload the stream directly to cloud storage (Azure Blob, Amazon S3, Google Cloud Storage) without creating a local file. Includes a mock IStreamProvider example and guidance for real cloud SDK integration.
// Keywords: Aspose.Cells | HTML to XLSX C# | MemoryStream upload | cloud storage upload | Azure Blob Storage | AWS S3 | Google Cloud Storage | IStreamProvider | in‑memory conversion | C# Excel export
// Common Searches: convert html table to excel c# aspose.cells | aspose.cells save workbook to memorystream | upload excel stream to azure blob using c# | aspose.cells upload xlsx to amazon s3 | c# html to xlsx and store in google cloud storage | how to use istreamprovider with aspose.cells
// Developer Intent: Create an Excel workbook from HTML and stream it directly to cloud storage without writing a temporary file.
// Use Cases: Generate server‑side reports from HTML and store them in Azure Blob for downstream BI pipelines. | Transform user‑submitted HTML forms into XLSX files and archive them in Amazon S3. | Produce Excel files from HTML content in a Cloud Function and write them to Google Cloud Storage.
// AI Prompts: Show how to replace the mock CloudStorageHelper with Azure Blob Storage SDK code for uploading the MemoryStream. | Provide an IStreamProvider implementation that writes Aspose.Cells output directly to an Amazon S3 stream using AWSSDK. | Explain how to modify the example to save the workbook as CSV while still uploading via the stream provider.

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsHtmlToExcel
{
    // Placeholder helper that represents uploading a stream to a cloud storage service.
    // In a real scenario replace the body of UploadAsync with actual SDK calls (e.g., Azure Blob, AWS S3, etc.).
    // Load an HTML file into an Aspose.Cells Workbook, save it as XLSX to a MemoryStream, and upload the stream directly to cloud storage (Azure Blob, Amazon S3, Google Cloud Storage) without creating a local file. Includes a mock IStreamProvider example and guidance for real cloud SDK integration.
    public static class CloudStorageHelper
    {
        public static async Task UploadAsync(Stream dataStream, string cloudPath)
        {
            // Ensure the stream is positioned at the beginning.
            dataStream.Position = 0;

            // Simulate async upload operation.
            await Task.Run(() =>
            {
                // For demonstration, write the stream to a local file that mimics cloud storage.
                // Replace this block with real cloud SDK upload logic.
                string localPath = Path.Combine(Path.GetTempPath(), cloudPath.Replace('/', Path.DirectorySeparatorChar));
                Directory.CreateDirectory(Path.GetDirectoryName(localPath));
                using (FileStream file = new FileStream(localPath, FileMode.Create, FileAccess.Write))
                {
                    dataStream.CopyTo(file);
                }
                Console.WriteLine($"[Mock Upload] Stream saved to local path: {localPath}");
            });
        }
    }

    // Example implementation of IStreamProvider that could be used when saving HTML.
    // Not required for the HTML‑to‑Excel conversion itself, but shown for completeness.
    public class HtmlExportStreamProvider : IStreamProvider
    {
        private readonly string _outputDirectory;

        public HtmlExportStreamProvider(string outputDirectory)
        {
            _outputDirectory = outputDirectory;
        }

        public void InitStream(StreamProviderOptions options)
        {
            // Determine the full path for the output file.
            string filePath = Path.Combine(_outputDirectory, options.CustomPath ?? options.DefaultPath);
            Directory.CreateDirectory(Path.GetDirectoryName(filePath));
            options.Stream = new FileStream(filePath, FileMode.Create, FileAccess.Write);
        }

        public void CloseStream(StreamProviderOptions options)
        {
            options.Stream?.Close();
        }
    }

    public class HtmlToExcelConverter
    {
        // Converts an HTML file to an Excel file and uploads the result directly to cloud storage.
        public static async Task ConvertAndUploadAsync(string htmlFilePath, string cloudDestinationPath)
        {
            // Load the HTML file into a workbook.
            // Aspose.Cells can load HTML directly via the Workbook constructor.
            Workbook workbook = new Workbook(htmlFilePath);

            // Save the workbook to a memory stream in XLSX format.
            using (MemoryStream excelStream = new MemoryStream())
            {
                workbook.Save(excelStream, SaveFormat.Xlsx);

                // Upload the generated Excel stream to cloud storage.
                await CloudStorageHelper.UploadAsync(excelStream, cloudDestinationPath);
            }

            // Optional: clean up the workbook instance.
            workbook.Dispose();
        }

        // Example usage.
        public static async Task RunDemoAsync()
        {
            string htmlInput = "sample.html";               // Path to the source HTML file.
            string cloudPath = "mycontainer/output.xlsx";   // Desired cloud storage path.

            // Ensure the HTML file exists for the demo.
            if (!File.Exists(htmlInput))
            {
                // Create a simple HTML file with a table for demonstration purposes.
                File.WriteAllText(htmlInput,
                    "<html><body><table><tr><th>Name</th><th>Age</th></tr>" +
                    "<tr><td>Alice</td><td>30</td></tr>" +
                    "<tr><td>Bob</td><td>25</td></tr></table></body></html>");
            }

            await ConvertAndUploadAsync(htmlInput, cloudPath);
        }
    }

    // Entry point for the console application.
    class Program
    {
        static async Task Main(string[] args)
        {
            await HtmlToExcelConverter.RunDemoAsync();
        }
    }
}
