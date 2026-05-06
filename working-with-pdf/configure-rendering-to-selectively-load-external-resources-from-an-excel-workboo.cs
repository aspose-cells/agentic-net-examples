using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

public class SelectiveResourceLoadingPdfConversion
{
    // Custom stream provider that skips loading resources marked as Skip,
    // otherwise loads them from the default file path.
    public class CustomStreamProvider : IStreamProvider
    {
        public void InitStream(StreamProviderOptions options)
        {
            // If the resource is marked to be skipped, provide a null stream.
            if (options.ResourceLoadingType == ResourceLoadingType.Skip)
            {
                options.Stream = Stream.Null;
            }
            else
            {
                // Load the resource normally from the file system.
                options.Stream = new FileStream(options.DefaultPath, FileMode.Open, FileAccess.Read);
            }
        }

        public void CloseStream(StreamProviderOptions options)
        {
            // Ensure the stream is properly closed.
            options.Stream?.Close();
        }
    }

    public static void Run()
    {
        // Path to the source Excel workbook.
        string sourcePath = "input.xlsx";

        // Load the workbook.
        Workbook workbook = new Workbook(sourcePath);

        // Assign the custom stream provider to control external resource loading.
        workbook.Settings.ResourceProvider = new CustomStreamProvider();

        // Configure PDF save options as needed.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Example: calculate formulas before conversion.
            CalculateFormula = true,
            // Example: embed standard Windows fonts.
            EmbedStandardWindowsFonts = true
        };

        // Save the workbook as PDF with selective resource loading.
        string outputPath = "output.pdf";
        workbook.Save(outputPath, pdfOptions);

        Console.WriteLine($"PDF conversion completed. Output saved to: {outputPath}");
    }
}

// Entry point for demonstration.
class Program
{
    static void Main()
    {
        SelectiveResourceLoadingPdfConversion.Run();
    }
}