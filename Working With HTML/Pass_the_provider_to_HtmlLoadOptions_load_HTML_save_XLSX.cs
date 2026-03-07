using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom stream provider used during HTML import
    public class HtmlLoadWithProviderDemo : IStreamProvider
    {
        // Called before a resource stream is requested
        public void InitStream(StreamProviderOptions options)
        {
            // No special initialization required for this demo
        }

        // Returns a stream for the requested resource (e.g., images, CSS)
        public Stream ProvideStream(StreamProviderOptions options)
        {
            // The DefaultPath property contains the full path of the resource file.
            // Open it for reading and return the stream.
            return File.OpenRead(options.DefaultPath);
        }

        // Called after the resource has been processed
        public void CloseStream(StreamProviderOptions options)
        {
            if (options.Stream != null)
            {
                options.Stream.Close();
                options.Stream = null;
            }
        }

        // Demonstrates loading an HTML file with the custom provider and saving as XLSX
        public static void Run()
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Create HtmlLoadOptions and assign the custom stream provider
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                StreamProvider = new HtmlLoadWithProviderDemo()
            };

            // Load the HTML file into a Workbook using the options
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Save the loaded workbook as an Excel file
            workbook.Save("output.xlsx");
        }

        // Entry point
        public static void Main(string[] args)
        {
            Run();
        }
    }
}