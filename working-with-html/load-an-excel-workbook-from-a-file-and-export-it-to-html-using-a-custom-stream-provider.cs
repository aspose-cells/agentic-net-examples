// Title: Export Excel to HTML with a Custom IStreamProvider in Aspose.Cells for .NET
// Description: Loads an Excel workbook, implements a custom IStreamProvider that writes each HTML part to a chosen folder, and saves the workbook as HTML using HtmlSaveOptions. All generated files are placed in the specified output directory.
// Keywords: Aspose.Cells | C# | IStreamProvider | HtmlSaveOptions | custom output folder | Excel to HTML export | stream provider example | save workbook as HTML | Aspose.Cells .NET | HTML export with custom directory
// Common Searches: Aspose.Cells custom IStreamProvider example | export Excel workbook to HTML in a specific folder | C# HtmlSaveOptions StreamProvider usage | how to control HTML output location with Aspose.Cells | save Excel as HTML with separate sheet files
// Developer Intent: Generate HTML from an Excel file while directing every HTML fragment and resource to a user‑defined directory via a custom stream provider.
// Use Cases: Create a self‑contained web folder for Excel reports that includes a main page, sheet‑specific pages, and assets. | Produce temporary HTML snapshots of Excel templates in a background service before further processing or uploading. | Integrate Aspose.Cells HTML export into a web API that must store each part in a secure, per‑user location.
// AI Prompts: Write a C# IStreamProvider that prefixes each exported HTML file with a timestamp. | Show how to modify the ExportStreamProvider to save CSS files in a subdirectory while keeping HTML files in the root output folder. | Provide error‑handling code for ExportStreamProvider when the target directory is read‑only or disk space is low.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsStreamProviderDemo
{
    // Custom stream provider that writes exported HTML parts to a specific directory
    // Loads an Excel workbook, implements a custom IStreamProvider that writes each HTML part to a chosen folder, and saves the workbook as HTML using HtmlSaveOptions. All generated files are placed in the specified output directory.
    public class ExportStreamProvider : IStreamProvider
    {
        private readonly string _outputDirectory;

        public ExportStreamProvider(string outputDirectory)
        {
            _outputDirectory = outputDirectory;
        }

        // Called by Aspose.Cells before writing each part (main HTML, sheet HTML, resources, etc.)
        public void InitStream(StreamProviderOptions options)
        {
            // Ensure the target directory exists
            Directory.CreateDirectory(_outputDirectory);

            // Use the default file name (e.g., sheet001.htm) and place it in the output directory
            string fileName = Path.GetFileName(options.DefaultPath);
            string fullPath = Path.Combine(_outputDirectory, fileName);

            // Set the custom path that will appear in the main HTML file
            options.CustomPath = fullPath;

            // Provide the stream where Aspose.Cells will write the content
            options.Stream = File.Create(fullPath);
        }

        // Called after the part has been written
        public void CloseStream(StreamProviderOptions options)
        {
            options.Stream?.Close();
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Path to the source Excel workbook
                string sourceFile = "input.xlsx";

                // Verify that the source file exists
                if (!File.Exists(sourceFile))
                {
                    Console.WriteLine($"Source file not found: {sourceFile}");
                    return;
                }

                // Load the workbook from the file system
                Workbook workbook = new Workbook(sourceFile);

                // Directory where all HTML files and resources will be saved
                string outputDir = Path.Combine(Path.GetTempPath(), "AsposeStreamProviderDemo");

                // Ensure the output directory exists before saving the main HTML file
                Directory.CreateDirectory(outputDir);

                // Configure HTML save options to use the custom stream provider
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    StreamProvider = new ExportStreamProvider(outputDir)
                };

                // Main HTML file path (the file that references other parts)
                string mainHtmlPath = Path.Combine(outputDir, "output.html");

                // Export the workbook to HTML using the configured options
                workbook.Save(mainHtmlPath, saveOptions);

                Console.WriteLine($"Workbook exported to HTML with custom stream provider at: {outputDir}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
