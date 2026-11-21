// Title: Export an Excel workbook to a temporary HTML file using Aspose.Cells in C# and clean up the file
// Common Searches: how to export an Excel workbook to HTML in a temporary folder with Aspose.Cells C# | using IFilePathProvider to obtain a temp file path for Aspose.Cells HTML export | remove temporary HTML file after saving a workbook with Aspose.Cells .NET | Aspose.Cells HtmlSaveOptions export only active worksheet with Base64 images | best practice for cleaning up temp files created by Aspose.Cells HTML export
// Tags: Aspose.Cells HTML export with temporary path | IFilePathProvider custom temp file provider | HtmlSaveOptions export active worksheet | Base64 image conversion in HTML export | post‑export file cleanup in .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Interface for providing temporary file paths
    public interface IFilePathProvider
    {
        // Returns a full file path in a temporary folder with the given extension (e.g., ".html")
        string GetTempFilePath(string extension);
    }

    // Simple implementation that uses the system temporary folder
    // The HtmlExporter class accepts an IFilePathProvider to generate a unique temporary .html file path, loads an Excel workbook, saves it as HTML using HtmlSaveOptions (exporting only the active worksheet and embedding images as Base64), returns the temporary file path, and provides a CleanUp method that safely deletes the temporary HTML file.
    public class TempFilePathProvider : IFilePathProvider
    {
        public string GetTempFilePath(string extension)
        {
            string fileName = Guid.NewGuid().ToString() + extension;
            return Path.Combine(Path.GetTempPath(), fileName);
        }
    }

    public class HtmlExporter
    {
        private readonly IFilePathProvider _pathProvider;

        public HtmlExporter(IFilePathProvider pathProvider)
        {
            _pathProvider = pathProvider ?? throw new ArgumentNullException(nameof(pathProvider));
        }

        // Exports the workbook at sourcePath to an HTML file stored in a temporary folder.
        // Returns the full path of the generated HTML file.
        public string ExportToHtml(string sourcePath)
        {
            if (string.IsNullOrEmpty(sourcePath))
                throw new ArgumentException("Source path must be provided.", nameof(sourcePath));

            // Load the workbook (using the lifecycle rule for loading)
            Workbook workbook = new Workbook(sourcePath);

            // Prepare HTML save options if needed
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true, // Export only the first worksheet; adjust as needed
                ExportImagesAsBase64 = true
            };

            // Get a temporary file path for the HTML output
            string tempHtmlPath = _pathProvider.GetTempFilePath(".html");

            // Save the workbook as HTML (using the lifecycle rule for saving)
            workbook.Save(tempHtmlPath, saveOptions);

            return tempHtmlPath;
        }

        // Cleans up the temporary HTML file after processing
        public void CleanUp(string tempFilePath)
        {
            if (string.IsNullOrEmpty(tempFilePath))
                return;

            try
            {
                if (File.Exists(tempFilePath))
                {
                    File.Delete(tempFilePath);
                }
            }
            catch (Exception ex)
            {
                // Log or handle the exception as appropriate for your application
                Console.Error.WriteLine($"Failed to delete temporary file '{tempFilePath}': {ex.Message}");
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main(string[] args)
        {
            // Assume args[0] contains the path to the source Excel file
            if (args.Length == 0)
            {
                Console.WriteLine("Please provide the path to the Excel file.");
                return;
            }

            string excelPath = args[0];
            IFilePathProvider pathProvider = new TempFilePathProvider();
            HtmlExporter exporter = new HtmlExporter(pathProvider);

            // Export to HTML
            string htmlPath = exporter.ExportToHtml(excelPath);
            Console.WriteLine($"HTML exported to temporary file: {htmlPath}");

            // Perform any processing needed with the HTML file here
            // For example, read its content:
            string htmlContent = File.ReadAllText(htmlPath);
            Console.WriteLine("HTML content length: " + htmlContent.Length);

            // Clean up the temporary file
            exporter.CleanUp(htmlPath);
            Console.WriteLine("Temporary file cleaned up.");
        }
    }
}
