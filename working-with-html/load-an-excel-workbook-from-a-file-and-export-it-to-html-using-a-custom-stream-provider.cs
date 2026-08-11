// Title: Export Excel to HTML with a Custom IStreamProvider in Aspose.Cells for .NET
// Description: Loads an Excel workbook, implements a custom IStreamProvider to create file streams for each HTML part, configures HtmlSaveOptions, and saves the workbook as HTML while ensuring the output folder exists.
// Keywords: Aspose.Cells | IStreamProvider | HtmlSaveOptions | C# export Excel to HTML | custom stream provider | .NET workbook.Save HTML | Excel to HTML with images and CSS | file stream creation | directory creation for HTML export | Aspose.Cells example
// Common Searches: Aspose.Cells custom IStreamProvider example | How to export Excel to HTML using HtmlSaveOptions | C# save workbook as HTML with custom stream | Create folder for Aspose.Cells HTML export | Export Excel workbook to multiple HTML files
// Developer Intent: Generate HTML output from an Excel workbook while controlling where each part (HTML, images, CSS) is written via a custom stream provider.
// Use Cases: Store HTML, images, and CSS in a predefined directory structure for web publishing. | Stream each HTML component directly to cloud storage (e.g., Azure Blob, AWS S3) instead of the local file system. | Create temporary HTML files in a background service, then zip and deliver them to end‑users.
// AI Prompts: Write a C# method that uses Aspose.Cells to export a workbook to HTML with an IStreamProvider that writes to MemoryStream objects. | Add robust error handling to ExportStreamProvider for missing permissions or invalid paths during HTML export. | Show how to modify ExportStreamProvider to embed CSS inline within the main HTML file rather than generating a separate stylesheet.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsStreamProviderDemo
{
    // Custom stream provider that creates a file stream for each exported part
    // Loads an Excel workbook, implements a custom IStreamProvider to create file streams for each HTML part, configures HtmlSaveOptions, and saves the workbook as HTML while ensuring the output folder exists.
    public class ExportStreamProvider : IStreamProvider
    {
        public void InitStream(StreamProviderOptions options)
        {
            // Use the default path supplied by Aspose.Cells
            string path = options.DefaultPath;

            // Ensure the directory exists
            Directory.CreateDirectory(Path.GetDirectoryName(path));

            // Create the file stream that Aspose.Cells will write to
            options.Stream = File.Create(path);
        }

        public void CloseStream(StreamProviderOptions options)
        {
            // Close the stream if it was created
            if (options.Stream != null)
            {
                options.Stream.Close();
            }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Load the workbook from the file
            Workbook workbook = new Workbook(sourcePath);

            // Set up HTML save options with the custom stream provider
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.StreamProvider = new ExportStreamProvider();

            // Destination HTML file (main file)
            string outputHtml = "output.html";

            // Save the workbook as HTML using the custom provider
            workbook.Save(outputHtml, saveOptions);

            Console.WriteLine($"Workbook successfully exported to HTML: {outputHtml}");
        }
    }
}
