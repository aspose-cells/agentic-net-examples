using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsStreamProviderDemo
{
    // Custom stream provider that supplies a read‑only stream for the HTML file.
    public class HtmlFileStreamProvider : IStreamProvider
    {
        // Called by Aspose.Cells before the stream is needed.
        public void InitStream(StreamProviderOptions options)
        {
            // The DefaultPath property contains the path of the HTML file Aspose.Cells wants to read.
            // Open the file for reading and assign it to the options.Stream.
            options.Stream = File.OpenRead(options.DefaultPath);
        }

        // Called after Aspose.Cells finishes using the stream.
        public void CloseStream(StreamProviderOptions options)
        {
            // Ensure the stream is properly closed and disposed.
            options.Stream?.Close();
            options.Stream = null;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source HTML file that represents a workbook.
            string htmlPath = "input.html";

            // Configure load options to use the custom stream provider.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                StreamProvider = new HtmlFileStreamProvider()
            };

            // Load the HTML document into a Workbook instance.
            Workbook workbook = new Workbook(htmlPath, loadOptions);

            // Export the workbook directly to XLSX format.
            string xlsxPath = "output.xlsx";
            workbook.Save(xlsxPath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML file '{htmlPath}' has been converted to '{xlsxPath}'.");
        }
    }
}