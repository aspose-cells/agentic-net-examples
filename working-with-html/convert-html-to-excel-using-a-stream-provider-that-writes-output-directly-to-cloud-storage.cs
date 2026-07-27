// Title: Convert HTML to Excel and stream directly to cloud storage with Aspose.Cells for .NET
// Description: Shows how to implement a custom IStreamProvider that creates a stream targeting cloud storage, load an HTML file into a Workbook using HtmlLoadOptions, and save the workbook as XLSX straight to the stream, avoiding intermediate local files.
// Keywords: Aspose.Cells | HTML to Excel | C# stream provider | IStreamProvider | cloud storage | Azure Blob Storage | Amazon S3 | save workbook to stream | convert HTML to XLSX | custom stream provider
// Common Searches: Aspose.Cells save workbook to cloud storage | C# convert HTML to XLSX using stream | How to use IStreamProvider with Aspose.Cells | Stream Aspose.Cells output to Azure Blob | Write Excel file directly to S3 from .NET
// Developer Intent: Write C# code that converts an HTML document into an Excel workbook and stores it directly in a cloud storage location via a custom stream provider.
// Use Cases: Generate Excel reports from HTML templates and store them in Azure Blob Storage without creating temporary files. | Expose a web API that receives HTML content, converts it to XLSX with Aspose.Cells, and streams the result to Amazon S3. | Batch‑process a directory of HTML files, converting each to XLSX and writing the outputs to a designated cloud folder using the same stream provider.
// AI Prompts: Provide a C# example that modifies CloudStorageStreamProvider to upload the generated XLSX to Azure Blob Storage instead of a local file. | Show how to make the stream provider asynchronous for efficient handling of large HTML files. | Explain how to configure HtmlLoadOptions to preserve CSS styles and merged cells when converting HTML to Excel with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Custom stream provider that writes directly to a cloud storage location (simulated here with a file path)
// Shows how to implement a custom IStreamProvider that creates a stream targeting cloud storage, load an HTML file into a Workbook using HtmlLoadOptions, and save the workbook as XLSX straight to the stream, avoiding intermediate local files.
public class CloudStorageStreamProvider : IStreamProvider
{
    private readonly string _cloudFilePath;

    public CloudStorageStreamProvider(string cloudFilePath)
    {
        _cloudFilePath = cloudFilePath;
    }

    // Initializes the stream that will be used for writing.
    public void InitStream(StreamProviderOptions options)
    {
        // Ensure the target directory exists.
        string directory = Path.GetDirectoryName(_cloudFilePath);
        if (!Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        // Create a file stream that represents the cloud storage destination.
        options.Stream = new FileStream(_cloudFilePath, FileMode.Create, FileAccess.Write);
        // CustomPath can be used by Aspose internally; set it to the file name.
        options.CustomPath = Path.GetFileName(_cloudFilePath);
    }

    // Closes the stream after the operation is complete.
    public void CloseStream(StreamProviderOptions options)
    {
        options.Stream?.Close();
    }
}

public class HtmlToExcelConverter
{
    public static void ConvertHtmlToExcel(string htmlFilePath, string cloudExcelPath)
    {
        // Load the HTML file into a workbook.
        // HtmlLoadOptions can be customized if needed; using defaults here.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        Workbook workbook = new Workbook(htmlFilePath, loadOptions);

        // Prepare the custom stream provider for cloud storage.
        CloudStorageStreamProvider streamProvider = new CloudStorageStreamProvider(cloudExcelPath);
        StreamProviderOptions providerOptions = new StreamProviderOptions();

        // Initialize the stream (creates the underlying file/stream).
        streamProvider.InitStream(providerOptions);

        // Save the workbook to the initialized stream in XLSX format.
        // This writes the Excel file directly to the cloud storage location.
        workbook.Save(providerOptions.Stream, SaveFormat.Xlsx);

        // Close and clean up the stream.
        streamProvider.CloseStream(providerOptions);
    }

    // Example usage.
    public static void Run()
    {
        // Path to the source HTML file.
        string sourceHtml = "input.html";

        // Destination path representing cloud storage (replace with actual cloud SDK path if needed).
        string cloudDestination = Path.Combine(Path.GetTempPath(), "CloudStorage", "output.xlsx");

        ConvertHtmlToExcel(sourceHtml, cloudDestination);

        Console.WriteLine($"HTML file '{sourceHtml}' has been converted and saved to cloud location '{cloudDestination}'.");
    }
}

// Entry point for demonstration.
class Program
{
    static void Main()
    {
        HtmlToExcelConverter.Run();
    }
}
