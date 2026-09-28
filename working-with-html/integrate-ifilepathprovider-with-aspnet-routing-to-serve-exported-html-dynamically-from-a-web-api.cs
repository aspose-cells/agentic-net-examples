// Title: Build an ASP.NET Core Web API that injects IFilePathProvider to locate Excel files and streams Aspose.Cells HTML export
// AI Prompts: Generate an ASP.NET Core controller with a GET route `/api/excel/{fileName}` that receives the file name, uses a dependency‑injected IFilePathProvider to resolve the path, calls ExportHelper.ExportToHtmlAsync, and returns the HTML via ContentResult with MIME type text/html. | Add service registration in Program.cs to bind IFilePathProvider to PhysicalFilePathProvider, reading the base directory from appsettings.json and validating the directory exists at startup. | Refactor ExportHelper so it writes the generated HTML directly to the HTTP response stream instead of returning a full string, reducing memory usage for large workbooks.
// Common Searches: how to return Aspose.Cells HTML export from ASP.NET Core Web API | inject custom file path provider for Excel to HTML conversion in .NET 6 | stream large Aspose.Cells HTML output without buffering in ASP.NET controller
// Tags: IFilePathProvider dependency injection ASP.NET Core | Aspose.Cells export workbook to HTML Web API | secure Excel file path resolution .NET | stream Aspose.Cells HTML response | configure base directory appsettings for file provider

using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

namespace MyApp
{
    // Simple interface to resolve file paths
    public interface IFilePathProvider
    {
        // Returns the absolute path for a given file name.
        string GetFilePath(string fileName);
    }

    // Physical implementation that combines a base directory with a safe file name
    // The example defines IFilePathProvider and its PhysicalFilePathProvider implementation for safe file‑path resolution, and ExportHelper.ExportToHtmlAsync which loads an Excel workbook with Aspose.Cells, saves it to HTML using HtmlSaveOptions (active worksheet only, images as Base64) and returns the HTML string. It can be integrated into an ASP.NET Core Web API to serve the generated HTML dynamically.
    public class PhysicalFilePathProvider : IFilePathProvider
    {
        private readonly string _baseDirectory;

        public PhysicalFilePathProvider(string baseDirectory)
        {
            _baseDirectory = baseDirectory ?? throw new ArgumentNullException(nameof(baseDirectory));
        }

        public string GetFilePath(string fileName)
        {
            // Prevent path traversal attacks
            var safeFileName = Path.GetFileName(fileName);
            return Path.Combine(_baseDirectory, safeFileName);
        }
    }

    public static class ExportHelper
    {
        // Exports the specified Excel file to HTML and returns the HTML string
        public static async Task<string> ExportToHtmlAsync(IFilePathProvider provider, string fileName)
        {
            var excelPath = provider.GetFilePath(fileName);

            if (!File.Exists(excelPath))
                throw new FileNotFoundException($"File '{fileName}' not found at path '{excelPath}'.");

            try
            {
                // Load the workbook using Aspose.Cells
                var workbook = new Workbook(excelPath);

                // Configure HTML save options
                var htmlOptions = new HtmlSaveOptions
                {
                    ExportActiveWorksheetOnly = true,
                    ExportImagesAsBase64 = true
                    // Note: HtmlSaveOptions does not have a Title property in current API versions.
                };

                // Save to a memory stream
                await using var htmlStream = new MemoryStream();
                workbook.Save(htmlStream, htmlOptions);
                htmlStream.Position = 0;

                // Read the generated HTML
                using var reader = new StreamReader(htmlStream);
                return await reader.ReadToEndAsync();
            }
            catch (Exception ex)
            {
                // Wrap the exception to provide context
                throw new InvalidOperationException($"Error processing file '{fileName}': {ex.Message}", ex);
            }
        }
    }

    class Program
    {
        // Entry point for a console application
        static async Task Main(string[] args)
        {
            // Expect a file name argument
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: dotnet run <ExcelFileName>");
                return;
            }

            var fileName = args[0];

            // Determine the directory where Excel files are stored (relative to the executable)
            var baseDirectory = Path.Combine(AppContext.BaseDirectory, "ExcelFiles");

            // Ensure the directory exists
            if (!Directory.Exists(baseDirectory))
            {
                Console.WriteLine($"Base directory '{baseDirectory}' does not exist.");
                return;
            }

            // Create the file path provider
            IFilePathProvider filePathProvider = new PhysicalFilePathProvider(baseDirectory);

            try
            {
                // Export to HTML
                var htmlContent = await ExportHelper.ExportToHtmlAsync(filePathProvider, fileName);

                // Output the HTML to console (or you could write to a file)
                Console.WriteLine(htmlContent);
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine(fnfEx.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
