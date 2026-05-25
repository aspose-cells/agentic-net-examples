using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Create a new workbook (empty or load an existing file)
Workbook workbook = new Workbook(); // or new Workbook("input.xlsx");

// Set a custom stream provider that skips loading external resources
workbook.Settings.ResourceProvider = new SkipResourceProvider();

// Optionally add some content to the workbook
Worksheet sheet = workbook.Worksheets[0];
sheet.Cells["A1"].PutValue("Sample PDF without external resources");

// Configure PDF save options as needed
PdfSaveOptions pdfOptions = new PdfSaveOptions
{
    // Example: calculate formulas before saving
    CalculateFormula = true,
    // Example: embed standard Windows fonts
    EmbedStandardWindowsFonts = true
};

// Save the workbook to PDF; external resources (e.g., linked images) will be skipped
workbook.Save("Output.pdf", pdfOptions);

// Custom IStreamProvider implementation that disables resource loading
class SkipResourceProvider : IStreamProvider
{
    public void InitStream(StreamProviderOptions options)
    {
        // Instruct Aspose.Cells to skip loading the linked resource
        options.ResourceLoadingType = ResourceLoadingType.Skip;
        // Provide a null stream to satisfy the API contract
        options.Stream = Stream.Null;
    }

    public void CloseStream(StreamProviderOptions options)
    {
        // No resources to close when skipping
        options.Stream?.Dispose();
    }
}