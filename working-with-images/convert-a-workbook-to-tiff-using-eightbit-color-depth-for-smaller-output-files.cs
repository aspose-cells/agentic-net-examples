// Title: Convert an Aspose.Cells Workbook to an 8‑bit TIFF with LZW compression (C#)
// Description: This C# sample demonstrates rendering a workbook to a TIFF file using Aspose.Cells, configuring the output to 8 bits per pixel and applying LZW compression to keep the file size low.
// Keywords: Aspose.Cells | C# TIFF export | 8‑bit TIFF | LZW compression | ImageOrPrintOptions | WorkbookRender | Excel to TIFF | low‑size image | ColorDepth.Format8bpp | TiffCompression.CompressionLZW
// Common Searches: Aspose.Cells export workbook to 8‑bit TIFF | C# render Excel as TIFF with LZW | set TIFF color depth 8bpp Aspose | reduce TIFF file size from Excel .NET | how to save Excel as TIFF using Aspose.Cells
// Developer Intent: Export a spreadsheet as a compact 8‑bit TIFF image.
// Use Cases: Create thumbnail previews of spreadsheets for web dashboards. | Archive Excel reports as small‑size TIFF files for compliance storage. | Batch‑convert multiple workbooks to 8‑bit TIFFs in a document‑management workflow.
// AI Prompts: Write C# code that loads an existing .xlsx file and saves it as an 8‑bit TIFF with LZW compression using Aspose.Cells. | Explain how to adjust DPI, page margins, and orientation while preserving 8‑bit color depth in the TIFF output. | Provide a loop that iterates through all worksheets in a workbook and creates separate 8‑bit TIFF files for each sheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// This C# sample demonstrates rendering a workbook to a TIFF file using Aspose.Cells, configuring the output to 8 bits per pixel and applying LZW compression to keep the file size low.
class ConvertWorkbookToTiff8bpp
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample data for 8‑bit TIFF conversion");

        // Configure rendering options for TIFF with 8‑bit color depth
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Tiff,               // Output format TIFF
            TiffColorDepth = ColorDepth.Format8bpp,   // 8‑bit per pixel
            TiffCompression = TiffCompression.CompressionLZW // Optional compression for smaller size
        };

        // Render the entire workbook to a TIFF file using the specified options
        WorkbookRender renderer = new WorkbookRender(workbook, options);
        renderer.ToImage("output_8bpp.tiff");

        // Release resources
        renderer.Dispose();

        Console.WriteLine("Workbook successfully converted to 8‑bit TIFF: output_8bpp.tiff");
    }
}
