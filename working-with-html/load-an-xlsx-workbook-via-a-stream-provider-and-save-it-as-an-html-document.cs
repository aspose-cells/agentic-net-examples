using System;
using System.IO;
using Aspose.Cells;

public class FileStreamProvider : IStreamProvider
{
    private readonly string _filePath;

    public FileStreamProvider(string filePath)
    {
        _filePath = filePath;
    }

    // Called by the consumer to obtain the stream.
    public void InitStream(StreamProviderOptions options)
    {
        // Open the file for reading and assign it to the options.
        options.Stream = File.OpenRead(_filePath);
    }

    // Called by the consumer to close the stream.
    public void CloseStream(StreamProviderOptions options)
    {
        if (options.Stream != null)
        {
            options.Stream.Close();
        }
    }
}

public class LoadAndSaveExample
{
    public static void Run()
    {
        // Path to the source XLSX file.
        string xlsxPath = "input.xlsx";

        // Path where the HTML output will be saved.
        string htmlPath = "output.html";

        // Create a stream provider that will supply a read‑only stream for the XLSX file.
        IStreamProvider provider = new FileStreamProvider(xlsxPath);

        // Prepare the options object that will receive the stream.
        StreamProviderOptions providerOptions = new StreamProviderOptions();

        // Initialise the stream (the provider sets providerOptions.Stream).
        provider.InitStream(providerOptions);

        // Load the workbook from the obtained stream.
        using (Stream xlsxStream = providerOptions.Stream)
        {
            Workbook workbook = new Workbook(xlsxStream);

            // Configure HTML save options (using the rule Save(string, SaveOptions)).
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Example: export only the active worksheet.
                ExportActiveWorksheetOnly = true
            };

            // Save the workbook as HTML.
            workbook.Save(htmlPath, htmlOptions);
        }

        // Close the stream via the provider (good practice).
        provider.CloseStream(providerOptions);
    }
}

// Entry point for demonstration.
class Program
{
    static void Main()
    {
        LoadAndSaveExample.Run();
        Console.WriteLine("Workbook loaded from stream and saved as HTML.");
    }
}