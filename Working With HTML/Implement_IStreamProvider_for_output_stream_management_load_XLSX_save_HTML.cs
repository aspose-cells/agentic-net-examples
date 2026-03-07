using System;
using System.IO;
using Aspose.Cells;

namespace AsposeStreamProviderDemo
{
    // Custom stream provider that writes all exported resources to a specific directory
    public class CustomStreamProvider : IStreamProvider
    {
        private readonly string _outputDirectory;

        public CustomStreamProvider(string outputDirectory)
        {
            _outputDirectory = outputDirectory;
        }

        // Called by Aspose.Cells when it needs a stream for a resource (HTML file, images, etc.)
        public void InitStream(StreamProviderOptions options)
        {
            // Determine the full path where the resource will be saved
            // If the caller provides a default path, use its file name inside the output directory
            string fileName = Path.GetFileName(options.DefaultPath);
            string fullPath = Path.Combine(_outputDirectory, fileName);

            // Ensure the directory exists
            Directory.CreateDirectory(Path.GetDirectoryName(fullPath));

            // Set the custom path (used in generated HTML) and the actual stream
            options.CustomPath = fileName;          // URL that will appear in the HTML
            options.Stream = File.Create(fullPath); // Physical file stream
        }

        // Called after the resource has been written
        public void CloseStream(StreamProviderOptions options)
        {
            if (options.Stream != null)
            {
                options.Stream.Close();
                options.Stream = null;
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare an output folder
            string outputDir = Path.Combine(Path.GetTempPath(), "AsposeStreamProviderDemo");
            Directory.CreateDirectory(outputDir);

            // Create a workbook with sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello StreamProvider");

            // Save the workbook as HTML using the custom stream provider
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.StreamProvider = new CustomStreamProvider(outputDir);
            string htmlPath = Path.Combine(outputDir, "output.html");
            workbook.Save(htmlPath, saveOptions);
            Console.WriteLine($"Workbook saved as HTML to: {htmlPath}");

            // Load the HTML back using the same stream provider (for any external resources)
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            loadOptions.StreamProvider = new CustomStreamProvider(outputDir);
            Workbook loadedWorkbook = new Workbook(htmlPath, loadOptions);

            // Verify that the data was loaded correctly
            string loadedValue = loadedWorkbook.Worksheets[0].Cells["A1"].StringValue;
            Console.WriteLine($"Loaded cell A1 value: {loadedValue}");
        }
    }
}