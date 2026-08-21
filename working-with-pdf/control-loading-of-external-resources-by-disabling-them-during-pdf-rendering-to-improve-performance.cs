// Title: Skip External Resources During PDF Rendering with Aspose.Cells for .NET
// Description: Demonstrates how to implement a custom IStreamProvider that sets ResourceLoadingType to Skip, assigns it to Workbook.Settings.ResourceProvider, and saves the workbook as PDF/A‑1b. The approach prevents linked images or other external files from being loaded, resulting in faster PDF generation and lower memory consumption.
// Keywords: Aspose.Cells PDF rendering performance | C# skip external resources | IStreamProvider ResourceLoadingType.Skip | disable linked images Aspose.Cells | PDF/A‑1b export .NET | custom resource provider Aspose.Cells
// Common Searches: how to prevent external images from loading in Aspose.Cells PDF export | custom IStreamProvider to skip resources Aspose.Cells | Aspose.Cells disable linked file loading during PDF conversion | improve PDF generation speed Aspose.Cells .NET
// Developer Intent: Avoid loading any external linked files while converting a workbook to PDF to reduce processing time and memory usage.
// Use Cases: Generate quick PDF previews of large spreadsheets when image files are unavailable. | Create server‑side PDF/A‑1b documents without requiring access to external media assets. | Lower memory footprint in batch PDF conversions that contain many linked pictures.
// AI Prompts: Show a C# example that uses IStreamProvider to skip external resources when saving a workbook to PDF with Aspose.Cells. | Explain the effect of setting Workbook.Settings.ResourceProvider to a custom provider on PDF rendering speed and resource handling. | Provide step‑by‑step guidance for combining PdfSaveOptions with a SkipResourceProvider for high‑performance PDF/A‑1b generation.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsPdfRendering
{
    // Custom stream provider that skips loading of external resources (e.g., linked images)
    // Demonstrates how to implement a custom IStreamProvider that sets ResourceLoadingType to Skip, assigns it to Workbook.Settings.ResourceProvider, and saves the workbook as PDF/A‑1b. The approach prevents linked images or other external files from being loaded, resulting in faster PDF generation and lower memory consumption.
    class SkipResourceProvider : IStreamProvider
    {
        // Called by Aspose.Cells when a linked resource needs to be accessed
        public void InitStream(StreamProviderOptions options)
        {
            // Instruct the renderer to skip the resource
            options.ResourceLoadingType = ResourceLoadingType.Skip;
            // Provide a null stream so no data is read
            options.Stream = Stream.Null;
        }

        // Called after the resource processing is finished
        public void CloseStream(StreamProviderOptions options)
        {
            // Dispose the stream if it exists
            options.Stream?.Dispose();
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Add some sample data and a linked picture to demonstrate the effect
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("PDF rendering with external resources disabled.");

                // Insert a picture that references an external file (LinkToFile) if it exists
                const string externalImagePath = "external_image.jpg";
                if (File.Exists(externalImagePath))
                {
                    int pictureIndex = sheet.Pictures.Add(2, 0, externalImagePath);
                    sheet.Pictures[pictureIndex].Placement = PlacementType.Move;
                }
                else
                {
                    Console.WriteLine($"External image '{externalImagePath}' not found. Skipping picture insertion.");
                }

                // Assign the custom stream provider to skip loading external resources
                workbook.Settings.ResourceProvider = new SkipResourceProvider();

                // Configure PDF save options (optional customizations)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Embed standard Windows fonts for better compatibility
                    EmbedStandardWindowsFonts = true,
                    // Set compliance level
                    Compliance = PdfCompliance.PdfA1b,
                    // Ignore rendering errors to avoid interruptions
                    IgnoreError = true
                };

                // Save the workbook as PDF; external resources will be skipped, improving performance
                const string outputPath = "output.pdf";
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"PDF saved with external resources disabled at '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
