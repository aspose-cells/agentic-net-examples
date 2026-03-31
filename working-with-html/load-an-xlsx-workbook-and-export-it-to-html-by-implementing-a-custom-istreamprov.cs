using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsStreamProviderDemo
{
    // Custom stream provider that writes each generated HTML part to a specified directory.
    public class CustomHtmlStreamProvider : IStreamProvider
    {
        private readonly string _outputDirectory;

        public CustomHtmlStreamProvider(string outputDirectory)
        {
            _outputDirectory = outputDirectory;
        }

        // Called by Aspose.Cells before writing a part (main HTML, sheet HTML, etc.).
        public void InitStream(StreamProviderOptions options)
        {
            // Ensure the target directory exists.
            Directory.CreateDirectory(_outputDirectory);

            // Extract the file name from the default path (e.g., "sheet001.htm").
            string fileName = Path.GetFileName(options.DefaultPath);

            // Build the full path where the part will be saved.
            string fullPath = Path.Combine(_outputDirectory, fileName);

            // Set CustomPath so that references inside the main HTML use the relative name.
            options.CustomPath = fileName;

            // Provide the stream that Aspose.Cells will write to.
            options.Stream = File.Create(fullPath);
        }

        // Called after the part has been written.
        public void CloseStream(StreamProviderOptions options)
        {
            options.Stream?.Close();
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source XLSX workbook.
            string inputPath = "input.xlsx";

            // Load the workbook from the file system.
            Workbook workbook = new Workbook(inputPath);

            // Directory where all HTML files (main file + sheet files) will be placed.
            string htmlOutputDir = Path.Combine(Directory.GetCurrentDirectory(), "HtmlOutput");

            // Path of the main HTML file (entry point).
            string mainHtmlPath = Path.Combine(htmlOutputDir, "workbook.html");

            // Configure HTML save options and attach the custom stream provider.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.StreamProvider = new CustomHtmlStreamProvider(htmlOutputDir);

            // Export the workbook to HTML using the configured options.
            workbook.Save(mainHtmlPath, saveOptions);

            Console.WriteLine($"Workbook successfully exported to HTML at: {mainHtmlPath}");
        }
    }
}