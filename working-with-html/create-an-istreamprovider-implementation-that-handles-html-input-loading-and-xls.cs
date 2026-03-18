using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsStreamProviderDemo
{
    // Custom stream provider that can supply streams for both loading HTML resources
    // and for saving workbook data to a memory stream.
    public class HtmlXlsxStreamProvider : IStreamProvider
    {
        // Called by Aspose.Cells when a stream is required.
        public void InitStream(StreamProviderOptions options)
        {
            // If the caller requests a user‑provided resource (e.g., an image referenced in HTML)
            // we give a fresh MemoryStream so the caller can read/write without touching the file system.
            if (options.ResourceLoadingType == ResourceLoadingType.UserProvided)
            {
                options.Stream = new MemoryStream();
                return;
            }

            // For normal loading of external resources, open the file indicated by DefaultPath.
            // This covers the case when Aspose.Cells loads linked files while importing HTML.
            if (!string.IsNullOrEmpty(options.DefaultPath))
            {
                // Ensure the file exists before opening; otherwise fall back to an empty stream.
                if (File.Exists(options.DefaultPath))
                {
                    options.Stream = File.OpenRead(options.DefaultPath);
                }
                else
                {
                    options.Stream = Stream.Null;
                }
            }
            else
            {
                // No path supplied – provide an empty stream.
                options.Stream = Stream.Null;
            }
        }

        // Called by Aspose.Cells when the stream is no longer needed.
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
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Load a workbook from an HTML file using the custom stream provider.
            // -----------------------------------------------------------------
            string htmlInputPath = "input.html"; // Path to the source HTML file.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions
            {
                // The provider will be used for any external resources referenced in the HTML.
                StreamProvider = new HtmlXlsxStreamProvider()
            };

            // Load the workbook. The provider will be invoked internally as needed.
            Workbook workbook = new Workbook(htmlInputPath, loadOptions);

            // -----------------------------------------------------------------
            // 2. Save the workbook to XLSX format using a memory stream.
            // -----------------------------------------------------------------
            // Prepare a StreamProviderOptions instance that holds the target stream.
            StreamProviderOptions saveOptions = new StreamProviderOptions
            {
                // Use a MemoryStream so the XLSX data stays in memory.
                Stream = new MemoryStream()
            };

            // Save the workbook into the provided stream.
            workbook.Save(saveOptions.Stream, SaveFormat.Xlsx);

            // At this point the MemoryStream contains the XLSX bytes.
            // Write them to a physical file (or use the stream directly as needed).
            string xlsxOutputPath = "output.xlsx";
            using (FileStream file = new FileStream(xlsxOutputPath, FileMode.Create, FileAccess.Write))
            {
                // Reset position to the beginning before copying.
                saveOptions.Stream.Position = 0;
                saveOptions.Stream.CopyTo(file);
            }

            // Clean up the stream via the provider (optional here because we dispose manually).
            new HtmlXlsxStreamProvider().CloseStream(saveOptions);

            Console.WriteLine($"HTML loaded from '{htmlInputPath}' and saved as XLSX to '{xlsxOutputPath}'.");
        }
    }
}