using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source Excel file (XLSX)
        string sourcePath = "input.xlsx";

        // Load the workbook from the Excel file
        Workbook workbook = new Workbook(sourcePath);

        // Create HTML save options and assign a custom stream provider
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.StreamProvider = new FileStreamProvider("HtmlOutput");

        // Save the workbook as HTML using the save options
        workbook.Save("output.html", htmlOptions);
    }
}

// Custom implementation of IStreamProvider for handling attached files (e.g., images)
class FileStreamProvider : IStreamProvider
{
    private readonly string _outputDirectory;

    public FileStreamProvider(string outputDirectory)
    {
        _outputDirectory = outputDirectory;
    }

    // Called by Aspose.Cells when a new stream is required
    public void InitStream(StreamProviderOptions options)
    {
        // Ensure the output directory exists
        Directory.CreateDirectory(_outputDirectory);

        // Build the full file path for the attached resource
        string fileName = Path.GetFileName(options.DefaultPath);
        string fullPath = Path.Combine(_outputDirectory, fileName);

        // Set the custom path and create the file stream
        options.CustomPath = fullPath;
        options.Stream = File.Create(fullPath);
    }

    // Called after the stream has been written
    public void CloseStream(StreamProviderOptions options)
    {
        options.Stream?.Close();
    }
}