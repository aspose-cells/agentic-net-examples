// Title: Export an Excel workbook to HTML with a custom FileStreamProvider to save images and CSS in a separate folder using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, creates a FileStreamProvider implementing IStreamProvider to write HTML resources to a specified folder, and saves the workbook as HTML. | Implement a custom stream provider class in C# that directs Aspose.Cells HTML export resources (images, CSS) to a user‑defined output directory.
// Common Searches: aspnet how to export Excel to HTML with images saved to a custom folder using Aspose.Cells | c# Aspose.Cells IStreamProvider example for HTML conversion | save Excel workbook as HTML and separate resource files with Aspose.Cells .NET | custom stream provider for Aspose.Cells HTML export tutorial | Aspose.Cells HtmlSaveOptions StreamProvider property usage in C#
// Tags: Aspose.Cells IStreamProvider HTML export | C# custom FileStreamProvider for Aspose.Cells | HtmlSaveOptions StreamProvider configuration | export Excel to HTML with resource folder | Aspose.Cells HTML resource handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHtmlExport
{
    // Custom stream provider to handle HTML resources (e.g., images, CSS) during export
    // The example loads an Excel file, defines a FileStreamProvider that writes HTML resources such as images and CSS to a designated output directory, configures HtmlSaveOptions with this provider, and saves the workbook as an HTML file using Aspose.Cells for .NET.
    public class FileStreamProvider : IStreamProvider
    {
        private readonly string _outputFolder;

        public FileStreamProvider(string outputFolder)
        {
            _outputFolder = outputFolder;
            // Ensure the output directory exists
            Directory.CreateDirectory(_outputFolder);
        }

        // Called once before any streams are requested
        public void InitStream(StreamProviderOptions options)
        {
            // No initialization required for this simple provider
        }

        // Called by Aspose.Cells when it needs a stream for a resource
        public Stream GetStream(string name, StreamProviderOptions options)
        {
            // Build full path for the resource file
            string filePath = Path.Combine(_outputFolder, name);
            // Create a FileStream for writing the resource (images, CSS, etc.)
            return new FileStream(filePath, FileMode.Create, FileAccess.Write);
        }

        // Called after the resource has been written; we can close or dispose if needed
        public void CloseStream(string name, StreamProviderOptions options)
        {
            // No additional actions required; streams are closed automatically by Aspose.Cells
        }

        // Required by IStreamProvider (overload without name)
        public void CloseStream(StreamProviderOptions options)
        {
            // No additional actions required
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the source Excel file
                string inputPath = "input.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook from the file system
                Workbook workbook = new Workbook(inputPath);

                // Configure HTML save options with the custom stream provider
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    StreamProvider = new FileStreamProvider("output_resources")
                };

                // Path for the main HTML file
                string htmlOutputPath = "output.html";

                // Save the workbook as HTML using the custom stream provider
                workbook.Save(htmlOutputPath, saveOptions);

                Console.WriteLine("Workbook exported to HTML successfully.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
