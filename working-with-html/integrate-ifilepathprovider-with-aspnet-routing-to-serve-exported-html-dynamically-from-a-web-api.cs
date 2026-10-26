// Title: Serve Excel workbook as HTML through an ASP.NET Core Web API using IFilePathProvider and Aspose.Cells
// Common Searches: how to return an Aspose.Cells HTML export from an ASP.NET Core controller | using IFilePathProvider to locate Excel files in an ASP.NET Web API endpoint | dynamic ASP.NET Core route that streams Excel-to-HTML conversion with Aspose.Cells | ASP.NET Core Web API method that returns HTML string generated from a workbook | configure HtmlSaveOptions for HTML5 output in an ASP.NET Core service
// Tags: asp.net core aspose.cells html export endpoint | IFilePathProvider workbook path resolution asp.net | excel to html conversion service asp.net core | memory stream html generation aspose.cells | routing dynamic html response asp.net core | htmlsaveoptions html5 configuration aspose.cells | serve workbook as html via web api

using System;
using System.IO;
using Aspose.Cells;

namespace MyAsposeApp
{
    // Interface to provide file paths based on a key (e.g., workbook identifier)
    public interface IFilePathProvider
    {
        string GetFilePath(string key);
    }

    // Simple implementation that maps a key to a .xlsx file under a base directory
    // The example shows how to integrate an IFilePathProvider with ASP.NET Core routing to expose a Web API that converts an Excel workbook to HTML on demand using Aspose.Cells. A provider maps a workbook key to a .xlsx file, the exporter loads the workbook, applies HtmlSaveOptions (active sheet only, base64 images, HTML5), writes the result to a MemoryStream, and the controller returns the HTML string as the response. This enables dynamic, file‑based Excel‑to‑HTML rendering through a single API route.
    public class FilePathProvider : IFilePathProvider
    {
        private readonly string _baseDirectory;

        public FilePathProvider(string baseDirectory)
        {
            _baseDirectory = baseDirectory ?? throw new ArgumentNullException(nameof(baseDirectory));
        }

        public string GetFilePath(string key)
        {
            var fileName = $"{key}.xlsx";
            var fullPath = Path.Combine(_baseDirectory, fileName);

            // Ensure the workbook file exists before returning the path
            if (!File.Exists(fullPath))
                throw new FileNotFoundException($"Workbook not found for key '{key}'.", fullPath);

            return fullPath;
        }
    }

    // Core exporter that converts a workbook to HTML
    public static class WorkbookHtmlExporter
    {
        public static string ExportToHtml(string workbookPath)
        {
            // Load the workbook using Aspose.Cells
            var workbook = new Workbook(workbookPath);

            // Configure HTML export options
            var htmlOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true, // Export only the active sheet
                ExportImagesAsBase64 = true,      // Embed images directly in HTML
                HtmlVersion = HtmlVersion.Html5   // Use HTML5 output
            };

            // Save the workbook to a memory stream in HTML format
            using var memoryStream = new MemoryStream();
            workbook.Save(memoryStream, htmlOptions);
            memoryStream.Position = 0;

            // Read the generated HTML string
            using var reader = new StreamReader(memoryStream);
            return reader.ReadToEnd();
        }
    }

    // Console application entry point
    internal class Program
    {
        private static void Main(string[] args)
        {
            // Expect a workbook identifier (key) as the first argument
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: dotnet run <workbook-key>");
                return;
            }

            string workbookKey = args[0];
            // Adjust the base directory to where your .xlsx files are stored
            string baseDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Workbooks");

            try
            {
                IFilePathProvider pathProvider = new FilePathProvider(baseDirectory);
                string workbookPath = pathProvider.GetFilePath(workbookKey);

                string htmlContent = WorkbookHtmlExporter.ExportToHtml(workbookPath);

                // Output the HTML to a file next to the executable
                string outputFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{workbookKey}.html");
                File.WriteAllText(outputFile, htmlContent);
                Console.WriteLine($"HTML exported successfully to: {outputFile}");
            }
            catch (FileNotFoundException ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}
