using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Custom stream provider to supply external resources (e.g., linked images) during PDF rendering
public class CustomStreamProvider : IStreamProvider
{
    // Called by Aspose.Cells when a resource stream is needed
    public void InitStream(StreamProviderOptions options)
    {
        // If the user wants to provide the stream manually
        if (options.ResourceLoadingType == ResourceLoadingType.UserProvided)
        {
            // Example: provide a stream for a linked image file
            // Adjust the path to point to an existing image on your system
            string imagePath = "linked_image.png";
            options.Stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read);
        }
        // If the resource should be skipped, provide a null stream
        else if (options.ResourceLoadingType == ResourceLoadingType.Skip)
        {
            options.Stream = Stream.Null;
        }
        // Otherwise load the resource normally from its default location
        else
        {
            // Default behavior: open the file specified in options.DefaultPath
            options.Stream = new FileStream(options.DefaultPath, FileMode.Open, FileAccess.Read);
        }
    }

    // Called after the resource has been processed
    public void CloseStream(StreamProviderOptions options)
    {
        options.Stream?.Close();
    }
}

public class LoadExternalResourceForPdfDemo
{
    public static void Run()
    {
        // Load an existing workbook that may contain external linked resources
        // Replace "input.xlsx" with the path to your workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Assign the custom stream provider to handle external resources during PDF conversion
        workbook.Settings.ResourceProvider = new CustomStreamProvider();

        // Create PDF save options (optional: configure additional settings here)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Example: embed standard Windows fonts
            EmbedStandardWindowsFonts = true,
            // Example: set PDF compliance level
            Compliance = PdfCompliance.PdfA1b,
            // Example: calculate formulas before rendering
            CalculateFormula = true
        };

        // Save the workbook as PDF; the custom stream provider will be invoked for linked resources
        workbook.Save("output.pdf", pdfOptions);
    }
}

// Entry point for demonstration
class Program
{
    static void Main()
    {
        LoadExternalResourceForPdfDemo.Run();
        Console.WriteLine("PDF generated with external resource handling.");
    }
}