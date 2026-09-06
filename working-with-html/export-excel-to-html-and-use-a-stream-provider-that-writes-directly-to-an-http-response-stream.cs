// Title: Export an Excel workbook to HTML with Aspose.Cells and stream the result directly to an HTTP response using a custom IStreamProvider (C#)
// AI Prompts: Write a C# method that loads a Workbook from a file path and saves it as HTML to any Stream by configuring HtmlSaveOptions with a custom stream provider. | Create a class that implements the Aspose.Cells IStreamProvider interface to return the same output Stream for all HTML resources such as images and CSS. | Develop a console program that receives an Excel file path and writes the generated HTML to the HttpResponse output stream via the custom stream provider.
// Common Searches: how to export Excel to HTML directly to HttpResponse stream using Aspose.Cells C# | Aspose.Cells custom IStreamProvider example for HTMLSaveOptions | C# write all Aspose.Cells HTML resources to a single output stream | stream Aspose.Cells HTML export to browser without creating temporary files
// Tags: Aspose.Cells HTMLSaveOptions custom stream provider | export workbook to HTML stream C# | single output stream for Aspose.Cells HTML export | write Aspose.Cells HTML output to HttpResponse | C# Excel to HTML conversion using IStreamProvider

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // Entry point for the console application.
    // The example demonstrates loading an Excel workbook with Aspose.Cells, configuring HtmlSaveOptions to use a custom IStreamProvider that directs every generated resource (HTML, images, CSS) to the same output Stream, and saving the workbook as HTML directly to that Stream. A console wrapper accepts input and output paths, enabling the HTML to be streamed to an HTTP response or any other Stream without intermediate files.
    public class Program
    {
        // Usage: AsposeCellsExample.exe <excelFilePath> <outputHtmlPath>
        public static void Main(string[] args)
        {
            try
            {
                if (args.Length < 2)
                {
                    Console.WriteLine("Please provide the Excel file path and the output HTML file path as arguments.");
                    return;
                }

                string excelFilePath = args[0];
                string outputHtmlPath = args[1];

                // Verify that the source Excel file exists.
                if (!File.Exists(excelFilePath))
                {
                    Console.WriteLine($"The Excel file '{excelFilePath}' was not found.");
                    return;
                }

                // Ensure the output directory exists.
                string outputDir = Path.GetDirectoryName(outputHtmlPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Export the Excel file to HTML.
                using (FileStream outputStream = new FileStream(outputHtmlPath, FileMode.Create, FileAccess.Write))
                {
                    var exporter = new ExcelToHtmlExporter();
                    exporter.Export(outputStream, excelFilePath);
                }

                Console.WriteLine($"Export completed successfully. HTML saved to '{outputHtmlPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }

    // Exports the Excel file located at excelFilePath to HTML and writes directly to the provided output stream.
    public class ExcelToHtmlExporter
    {
        public void Export(Stream outputStream, string excelFilePath)
        {
            try
            {
                // Load the workbook from the specified file.
                var workbook = new Workbook(excelFilePath);

                // Configure HTML save options.
                var htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    // Direct all generated resources (images, CSS) to the same output stream.
                    StreamProvider = new OutputStreamProvider(outputStream)
                };

                // Save the workbook directly to the output stream.
                workbook.Save(outputStream, htmlOptions);
            }
            catch (Exception ex)
            {
                // In case of error, write the message to the output stream as plain text.
                using (var writer = new StreamWriter(outputStream, leaveOpen: true))
                {
                    writer.Write($"Error exporting Excel to HTML: {ex.Message}");
                    writer.Flush();
                }
            }
        }
    }

    // Custom stream provider that directs all resource streams to a single output stream.
    public class OutputStreamProvider : IStreamProvider
    {
        private readonly Stream _outputStream;

        public OutputStreamProvider(Stream outputStream)
        {
            _outputStream = outputStream;
        }

        // Returns the output stream for a given resource name.
        public Stream GetStream(string name) => _outputStream;

        // Overload that also receives the file extension; still returns the output stream.
        public Stream GetStream(string name, string extension) => _outputStream;

        // Called before any streams are requested; no initialization needed for the output stream.
        public void InitStream(StreamProviderOptions options)
        {
            // No action required.
        }

        // Called after the stream is no longer needed; no action required for the output stream.
        public void CloseStream(string name, Stream stream)
        {
            // No action required.
        }

        // Overload with options (required by newer IStreamProvider definitions).
        public void CloseStream(string name, Stream stream, StreamProviderOptions options)
        {
            // No action required.
        }

        // Additional overload that may be required by some versions of the interface.
        public void CloseStream(StreamProviderOptions options)
        {
            // No action required.
        }
    }
}
