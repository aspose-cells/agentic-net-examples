// Title: Export an Excel workbook to HTML in C# with Aspose.Cells using a custom IStreamProvider for direct HttpResponse streaming
// Common Searches: c# aspnet export excel to html using aspose.cells streamprovider | aspose.cells write html output to httpresponse stream | save workbook as html with base64 images using aspose.cells c# | custom iStreamProvider implementation for aspose.cells html export | aspose.cells htmlsaveoptions exportactiveworksheetonly example
// Tags: Aspose.Cells IStreamProvider custom implementation | Aspose.Cells HTMLSaveOptions ExportActiveWorksheetOnly | Aspose.Cells export workbook to HTML with Base64 images | C# stream HTML result to HttpResponse using Aspose.Cells | Aspose.Cells HTML5 UTF-8 conversion

using System;
using System.IO;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to load or create a Workbook, configure HtmlSaveOptions (HTML5, UTF‑8, ExportActiveWorksheetOnly, Base64 images), and save the workbook as HTML using a custom IStreamProvider that forwards all resource streams to a supplied target stream, enabling direct streaming of the generated HTML to an HTTP response in a C# web application.
public class ExcelToHtmlHandler
{
    // Supplies a target stream to Aspose.Cells for HTML output.
    private class StreamProvider : IStreamProvider
    {
        private readonly Stream _targetStream;

        public StreamProvider(Stream targetStream)
        {
            _targetStream = targetStream ?? throw new ArgumentNullException(nameof(targetStream));
        }

        public void InitStream(StreamProviderOptions options) { }

        // Returns the same stream for all resources (images are embedded as Base64).
        public Stream GetStream(string name, out bool isNew)
        {
            isNew = true;
            return _targetStream;
        }

        public void CloseStream(StreamProviderOptions options) { }
    }

    /// <param name="inputPath">Path to the source .xlsx file.</param>
    /// <param name="outputPath">Path where the generated HTML will be saved.</param>
    public void ConvertToHtml(string inputPath, string outputPath)
    {
        try
        {
            // Load existing workbook or create a new empty one.
            Workbook workbook;
            if (!string.IsNullOrEmpty(inputPath) && File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
            }

            // Configure HTML save options.
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true,
                ExportImagesAsBase64 = true,
                HtmlVersion = HtmlVersion.Html5,
                Encoding = Encoding.UTF8
            };

            // Ensure output directory exists.
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Open a file stream for the HTML output and assign the custom stream provider.
            using (FileStream fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.Read))
            {
                htmlOptions.StreamProvider = new StreamProvider(fs);
                // Save workbook as HTML using the specified options.
                workbook.Save(outputPath, htmlOptions);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error converting Excel to HTML: {ex.Message}");
            throw;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            string input = args.Length > 0 ? args[0] : "input.xlsx";
            string output = args.Length > 1 ? args[1] : "output.html";

            var handler = new ExcelToHtmlHandler();
            handler.ConvertToHtml(input, output);

            Console.WriteLine($"Conversion completed. HTML saved to '{output}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
