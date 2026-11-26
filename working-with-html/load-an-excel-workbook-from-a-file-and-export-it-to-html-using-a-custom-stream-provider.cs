// Title: Export an Excel workbook to HTML using a custom IStreamProvider to store images and sheet fragments in a designated folder (C# Aspose.Cells)
// Common Searches: how to implement a custom IStreamProvider for HTML export in Aspose.Cells C# | save Excel workbook as HTML with external resources placed in a specific directory using Aspose.Cells | Aspose.Cells C# example for directing image files to a custom folder during HTML conversion | C# code to export Excel to HTML and control the path of generated HTML fragments with IStreamProvider | Aspose.Cells HTMLSaveOptions custom resource path tutorial
// Tags: custom IStreamProvider HTML export Aspose.Cells | Aspose.Cells HTMLSaveOptions external resource directory | C# export workbook to HTML with custom resource path | Aspose.Cells control image output folder | HTML conversion external files handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Custom stream provider that writes external resources (e.g., images, sheet html fragments)
    // to a specified directory and sets a custom relative path used in the generated HTML.
    // The sample loads an Excel workbook, configures HtmlSaveOptions to export the entire workbook to HTML, and attaches a CustomHtmlStreamProvider that creates a specified folder, generates custom relative URLs, and writes external resources (images, sheet fragments) to that folder while saving the main HTML file.
    public class CustomHtmlStreamProvider : IStreamProvider
    {
        private readonly string _resourceDirectory;

        public CustomHtmlStreamProvider(string resourceDirectory)
        {
            _resourceDirectory = resourceDirectory;
        }

        // Called by Aspose.Cells before a resource stream is needed.
        public void InitStream(StreamProviderOptions options)
        {
            // Ensure the target directory exists.
            Directory.CreateDirectory(_resourceDirectory);

            // Derive a file name from the default path supplied by the library.
            // DefaultPath is something like "output_files/sheet001.htm".
            string fileName = Path.GetFileName(options.DefaultPath);
            if (string.IsNullOrEmpty(fileName))
            {
                // Fallback to a GUID if the library does not provide a name.
                fileName = Guid.NewGuid().ToString() + ".dat";
            }

            // Set the custom URL that will appear in the main HTML file.
            // Use forward slashes for URL compatibility.
            options.CustomPath = $"{_resourceDirectory.Replace("\\", "/")}/{fileName}";

            // Create the physical file stream where the resource will be written.
            string fullPath = Path.Combine(_resourceDirectory, fileName);
            options.Stream = new FileStream(fullPath, FileMode.Create, FileAccess.Write);
        }

        // Called after the resource has been written.
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
        static void Main(string[] args)
        {
            // Paths – adjust as needed. The input workbook must exist.
            string inputWorkbookPath = "input.xlsx";
            string outputHtmlPath = "output.html";
            string resourceDirectory = "output_resources";

            try
            {
                // Load the workbook from the file system.
                Workbook workbook = new Workbook(inputWorkbookPath);

                // Configure HTML save options.
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export the whole workbook (set to true to export only the active sheet).
                    ExportActiveWorksheetOnly = false,

                    // Attach the custom stream provider so that external resources are saved
                    // into the specified folder and referenced via a custom URL.
                    StreamProvider = new CustomHtmlStreamProvider(resourceDirectory)
                };

                // Perform the conversion.
                workbook.Save(outputHtmlPath, htmlOptions);

                Console.WriteLine($"Workbook '{inputWorkbookPath}' successfully exported to HTML:");
                Console.WriteLine($"  Main HTML file: {Path.GetFullPath(outputHtmlPath)}");
                Console.WriteLine($"  External resources directory: {Path.GetFullPath(resourceDirectory)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred during the export process:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
