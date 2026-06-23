using System;
using System.IO;
using Aspose.Cells;

// Custom stream provider that writes exported HTML parts to a specified directory
class CustomStreamProvider : IStreamProvider
{
    private readonly string _outputDirectory;

    public CustomStreamProvider(string outputDirectory)
    {
        _outputDirectory = outputDirectory;
    }

    // Called by Aspose.Cells when a new stream is required
    public void InitStream(StreamProviderOptions options)
    {
        // Use the default file name (e.g., sheet001.htm) and place it in the output directory
        string fileName = Path.GetFileName(options.DefaultPath);
        string fullPath = Path.Combine(_outputDirectory, fileName);

        // Ensure the directory exists
        Directory.CreateDirectory(Path.GetDirectoryName(fullPath));

        // Set the custom path that will be referenced from the main HTML file
        options.CustomPath = fullPath;

        // Create the file stream for writing
        options.Stream = File.Create(fullPath);
    }

    // Called after the stream has been used
    public void CloseStream(StreamProviderOptions options)
    {
        if (options.Stream != null)
        {
            options.Stream.Close();
        }
    }
}

class Program
{
    static void Main()
    {
        // Path to the source Excel workbook
        string sourceExcelPath = "input.xlsx";

        // Directory where the HTML file and its related resources will be saved
        string outputDirectory = Path.Combine(Path.GetTempPath(), "AsposeHtmlExport");
        Directory.CreateDirectory(outputDirectory);

        // Load the workbook from the file (uses the Workbook(string) constructor)
        Workbook workbook = new Workbook(sourceExcelPath);

        // Configure HTML save options and assign the custom stream provider
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.StreamProvider = new CustomStreamProvider(outputDirectory);

        // Destination HTML file path
        string htmlOutputPath = Path.Combine(outputDirectory, "output.html");

        // Save the workbook as HTML using the options (uses Workbook.Save(string, SaveOptions))
        workbook.Save(htmlOutputPath, htmlOptions);

        Console.WriteLine($"Workbook successfully exported to HTML at: {htmlOutputPath}");
    }
}