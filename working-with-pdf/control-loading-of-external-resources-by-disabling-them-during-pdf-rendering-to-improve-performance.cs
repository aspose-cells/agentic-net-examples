// Title: C# – Disable External Resource Loading When Converting Excel to PDF with Aspose.Cells
// Description: Shows how to stop Aspose.Cells from loading linked images, charts, or data connections during PDF export by using a custom IStreamProvider that sets ResourceLoadingType to Skip, thereby boosting conversion speed and lowering memory usage.
// Keywords: Aspose.Cells PDF conversion | disable external resources | ResourceLoadingType.Skip | IStreamProvider C# | skip linked images Aspose | Excel to PDF performance | Aspose.Cells .NET | skip external data connections | PDF rendering optimization | custom stream provider
// Common Searches: Aspose.Cells skip external images PDF | How to disable resource loading in Aspose.Cells PDF export | C# custom IStreamProvider for PDF conversion | Improve PDF conversion speed Aspose.Cells | ResourceLoadingType.Skip example
// Developer Intent: Generate a PDF from an Excel workbook without loading any external resources to achieve faster, lighter conversions.
// Use Cases: Convert large workbooks that contain many linked pictures to PDF while omitting the pictures to reduce memory footprint. | Produce PDF reports on a server where external data connections are blocked, ensuring the conversion completes quickly. | Run an automated batch job that converts dozens of spreadsheets to PDF, consistently skipping all external assets for predictable performance.
// AI Prompts: Write a C# snippet that uses Aspose.Cells to save a workbook as PDF while skipping all external resources via a custom IStreamProvider. | Explain the effect of ResourceLoadingType.Skip on the PDF rendering pipeline in Aspose.Cells and compare it with other ResourceLoadingType options. | Show how to modify SkipResourceProvider to log the name of each skipped resource instead of returning Stream.Null.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to stop Aspose.Cells from loading linked images, charts, or data connections during PDF export by using a custom IStreamProvider that sets ResourceLoadingType to Skip, thereby boosting conversion speed and lowering memory usage.
class DisableExternalResourcesPdf
{
    static void Main()
    {
        // Load the workbook (replace with your source file)
        Workbook workbook = new Workbook("input.xlsx");

        // Assign a custom stream provider that skips loading external resources
        workbook.Settings.ResourceProvider = new SkipResourceProvider();

        // Configure PDF save options as needed
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            CalculateFormula = true // optional: calculate formulas before rendering
        };

        // Save the workbook as PDF; external resources will be ignored
        workbook.Save("output.pdf", pdfOptions);
    }
}

// Custom IStreamProvider implementation that disables resource loading
class SkipResourceProvider : IStreamProvider
{
    public void InitStream(StreamProviderOptions options)
    {
        // Instruct the renderer to skip the linked resource
        options.ResourceLoadingType = ResourceLoadingType.Skip;
        options.Stream = Stream.Null; // No stream needed when skipping
    }

    public void CloseStream(StreamProviderOptions options)
    {
        // No action required; ensure any non-null stream is disposed safely
        if (options.Stream != null && options.Stream != Stream.Null)
        {
            options.Stream.Dispose();
        }
    }
}
