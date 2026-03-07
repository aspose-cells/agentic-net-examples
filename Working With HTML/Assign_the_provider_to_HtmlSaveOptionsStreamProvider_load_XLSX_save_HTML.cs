using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsStreamProviderDemo
{
    // Custom stream provider that writes exported HTML resources to a specified directory.
    public class ExportStreamProvider : IStreamProvider
    {
        private readonly string _outputDirectory;

        public ExportStreamProvider(string outputDirectory)
        {
            _outputDirectory = outputDirectory;
        }

        // Called by Aspose.Cells before a resource stream is needed.
        public void InitStream(StreamProviderOptions options)
        {
            // Determine the full file path for the resource.
            // If the user supplied a custom path, use it; otherwise fall back to the default path.
            string filePath = Path.Combine(_outputDirectory,
                string.IsNullOrEmpty(options.CustomPath) ? options.DefaultPath : options.CustomPath);

            // Ensure the directory exists.
            Directory.CreateDirectory(Path.GetDirectoryName(filePath));

            // Assign the stream that Aspose.Cells will write to.
            options.Stream = new FileStream(filePath, FileMode.Create, FileAccess.Write);
        }

        // Called by Aspose.Cells after the resource has been written.
        public void CloseStream(StreamProviderOptions options)
        {
            options.Stream?.Close();
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the source XLSX file.
            string inputPath = "input.xlsx";

            // Load the workbook from the XLSX file.
            Workbook workbook = new Workbook(inputPath);

            // Directory where HTML and its related resources will be saved.
            string outputDir = Path.Combine(Path.GetTempPath(), "AsposeHtmlExport");
            Directory.CreateDirectory(outputDir);

            // Configure HTML save options with the custom stream provider.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.StreamProvider = new ExportStreamProvider(outputDir);

            // Save the workbook as HTML. The main HTML file is saved in the output directory.
            string htmlPath = Path.Combine(outputDir, "output.html");
            workbook.Save(htmlPath, htmlOptions);

            Console.WriteLine($"Workbook saved as HTML to: {htmlPath}");
        }
    }
}