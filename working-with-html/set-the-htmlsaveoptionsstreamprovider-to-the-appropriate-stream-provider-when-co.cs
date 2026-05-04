using System;
using System.IO;
using Aspose.Cells;

class HtmlExportWithStreamProvider
{
    static void Main()
    {
        // Create a temporary directory for the output files
        string outputDir = Path.Combine(Path.GetTempPath(), "AsposeHtmlExport");
        Directory.CreateDirectory(outputDir);

        // Create a workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("StreamProvider Demo");

        // Configure HTML save options with a custom stream provider
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.StreamProvider = new ExportStreamProvider(outputDir);

        // Save the workbook as HTML using the configured options
        string htmlPath = Path.Combine(outputDir, "output.html");
        workbook.Save(htmlPath, saveOptions);

        Console.WriteLine($"HTML file saved to: {htmlPath}");
    }

    // Custom implementation of IStreamProvider for exporting HTML resources
    class ExportStreamProvider : IStreamProvider
    {
        private readonly string _outputDirectory;

        public ExportStreamProvider(string outputDirectory)
        {
            _outputDirectory = outputDirectory;
        }

        // Called by Aspose.Cells before writing a resource
        public void InitStream(StreamProviderOptions options)
        {
            // Determine the file path: use CustomPath if provided, otherwise DefaultPath
            string relativePath = options.CustomPath ?? options.DefaultPath;
            string fullPath = Path.Combine(_outputDirectory, relativePath);

            // Ensure the target directory exists
            Directory.CreateDirectory(Path.GetDirectoryName(fullPath));

            // Assign a writable file stream to the options
            options.Stream = new FileStream(fullPath, FileMode.Create, FileAccess.Write);
        }

        // Called after the resource has been written
        public void CloseStream(StreamProviderOptions options)
        {
            // Close the stream if it was created
            options.Stream?.Close();
        }
    }
}