using System;
using System.IO;
using Aspose.Cells;

namespace AsposeStreamProviderDemo
{
    // Custom stream provider that creates output files in a specified directory
    public class CustomStreamProvider : IStreamProvider
    {
        private readonly string _outputDirectory;

        public CustomStreamProvider(string outputDirectory)
        {
            _outputDirectory = outputDirectory;
        }

        // Called by Aspose.Cells before writing each stream (e.g., main HTML file, images, etc.)
        public void InitStream(StreamProviderOptions options)
        {
            // Ensure the file name is just the file name part of the default path
            string fileName = Path.GetFileName(options.DefaultPath);
            options.CustomPath = fileName;

            // Build the full path inside the output directory
            string fullPath = Path.Combine(_outputDirectory, fileName);
            Directory.CreateDirectory(Path.GetDirectoryName(fullPath));

            // Create the file stream that Aspose.Cells will write to
            options.Stream = File.Create(fullPath);
        }

        // Called after the stream has been written
        public void CloseStream(StreamProviderOptions options)
        {
            if (options.Stream != null)
            {
                options.Stream.Close();
                options.Stream = null;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source XLSX workbook
            string sourcePath = "sample.xlsx";

            // Directory where all HTML related files will be placed
            string outputDir = Path.Combine(Path.GetTempPath(), "AsposeStreamDemo");
            Directory.CreateDirectory(outputDir);

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(sourcePath);

            // Configure HTML save options with the custom stream provider
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.StreamProvider = new CustomStreamProvider(outputDir);

            // Define the main HTML file name (the provider will place it in outputDir)
            string htmlFilePath = Path.Combine(outputDir, "output.html");

            // Save the workbook as HTML using the custom stream provider
            workbook.Save(htmlFilePath, saveOptions);

            Console.WriteLine($"Workbook exported to HTML at: {htmlFilePath}");
        }
    }
}