using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Path to the source Excel workbook
        string sourcePath = "input.xlsx";

        // Load the workbook from the file system
        Workbook workbook = new Workbook(sourcePath);

        // Configure image rendering options for high‑quality TIFF output
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Tiff,                     // Output format
            TiffCompression = TiffCompression.CompressionLZW, // Lossless compression
            HorizontalResolution = 300,                     // DPI for better clarity
            VerticalResolution = 300,
            OnePagePerSheet = true                          // Preserve original page layout
        };

        // Create a renderer for the entire workbook
        WorkbookRender renderer = new WorkbookRender(workbook, options);

        // Destination path for the TIFF image
        string outputPath = "output.tiff";

        // Render the whole workbook to a multi‑page TIFF file
        renderer.ToImage(outputPath);

        Console.WriteLine($"Workbook successfully rendered to TIFF: {outputPath}");
    }
}