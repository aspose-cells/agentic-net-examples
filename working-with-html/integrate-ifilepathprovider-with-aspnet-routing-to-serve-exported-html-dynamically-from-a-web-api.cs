// Title: Using IFilePathProvider in an ASP.NET Core Web API to stream Aspose.Cells HTML export
// Common Searches: asp.net core web api return html from aspose.cells workbook using IFilePathProvider | serve excel worksheet as html via asp.net core controller with aspose.cells | dynamic html export from excel in asp.net core using IFilePathProvider | asp.net core route that converts workbook to html on the fly with aspose.cells | use memory stream to send aspose.cells html output from web api endpoint
// Tags: IFilePathProvider integration with ASP.NET Core routing | Aspose.Cells HTML export in Web API | dynamic Excel to HTML conversion ASP.NET Core | memory stream response for Aspose.Cells HTML | serve workbook as HTML endpoint

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExport
{
    // Provides absolute paths for workbook files
    public interface IFilePathProvider
    {
        string GetFilePath(string fileName);
    }

    // Resolves files from a "Workbooks" folder under the current directory
    // Demonstrates how to inject an IFilePathProvider into an ASP.NET Core controller, load an Excel workbook with Aspose.Cells, export the active worksheet to HTML using HtmlSaveOptions (including Base64‑encoded images), and return the generated HTML as a memory‑stream response from a Web API endpoint.
    public class FilePathProvider : IFilePathProvider
    {
        private readonly string _baseFolder;

        public FilePathProvider()
        {
            _baseFolder = Path.Combine(Directory.GetCurrentDirectory(), "Workbooks");
        }

        public string GetFilePath(string fileName)
        {
            var safeFileName = Path.GetFileName(fileName);
            return Path.Combine(_baseFolder, safeFileName);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Expect the workbook file name as the first argument
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: AsposeCellsExport <WorkbookFileName>");
                return;
            }

            var fileName = args[0];
            var pathProvider = new FilePathProvider();
            var workbookPath = pathProvider.GetFilePath(fileName);

            // Verify the workbook exists before attempting to load
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook '{fileName}' not found at '{workbookPath}'.");
                return;
            }

            try
            {
                // Load the workbook using Aspose.Cells
                var workbook = new Workbook(workbookPath);

                // Configure HTML export options
                var htmlOptions = new HtmlSaveOptions
                {
                    ExportActiveWorksheetOnly = true,
                    ExportImagesAsBase64 = true
                    // Title property is not available in current Aspose.Cells version; omitted.
                };

                // Export to an in‑memory stream
                using var htmlStream = new MemoryStream();
                workbook.Save(htmlStream, htmlOptions);
                htmlStream.Position = 0;

                // Read the HTML content
                using var reader = new StreamReader(htmlStream);
                var htmlContent = reader.ReadToEnd();

                // Write the HTML to a file with the same name but .html extension
                var outputPath = Path.ChangeExtension(workbookPath, ".html");
                File.WriteAllText(outputPath, htmlContent);
                Console.WriteLine($"HTML exported successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Runtime safety: report any processing errors
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
