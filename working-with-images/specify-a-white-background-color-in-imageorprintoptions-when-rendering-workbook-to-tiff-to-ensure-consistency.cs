// Title: Aspose.Cells for .NET – Convert Excel Worksheet to TIFF with Solid White Background (C#)
// Description: This C# example creates a workbook, adds sample data, and uses ImageOrPrintOptions to export the sheet as a TIFF file. Setting Transparent = false produces a solid white canvas; the sample also shows 24‑bit color depth and LZW compression.
// Keywords: Aspose.Cells | C# | .NET | Excel to TIFF | white background | ImageOrPrintOptions | Transparent false | 24‑bit color | LZW compression | WorkbookRender | image export
// Common Searches: Aspose.Cells export Excel to TIFF with white background | C# set Transparent false for TIFF rendering | How to disable transparency in Aspose.Cells TIFF output | Set color depth and compression for TIFF in Aspose.Cells | Render workbook as TIFF using ImageOrPrintOptions
// Developer Intent: Produce a TIFF image from an Excel workbook that contains a non‑transparent white background.
// Use Cases: Generate print‑ready TIFF files for reports where a non‑transparent background is required. | Archive Excel worksheets as high‑color‑depth TIFFs with LZW compression for long‑term storage. | Create image assets of charts for inclusion in PDFs or web pages with consistent background color.
// AI Prompts: Write C# code that uses Aspose.Cells to save an Excel file as a TIFF image with a solid white background, 24‑bit depth, and LZW compression. | Explain the effect of ImageOrPrintOptions.Transparent = false on TIFF rendering in Aspose.Cells. | Provide a reusable function that accepts an Excel path and outputs a TIFF with configurable compression and a white background.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// This C# example creates a workbook, adds sample data, and uses ImageOrPrintOptions to export the sheet as a TIFF file. Setting Transparent = false produces a solid white canvas; the sample also shows 24‑bit color depth and LZW compression.
class RenderTiffWithWhiteBackground
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook(); // create
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("White background test");

        // Configure rendering options for TIFF
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Tiff,          // output format
            Transparent = false,                 // enforce white background
            TiffColorDepth = ColorDepth.Format24bpp, // optional: high color depth
            TiffCompression = TiffCompression.CompressionLZW // optional: compression
        };

        // Render the workbook to a TIFF file
        WorkbookRender renderer = new WorkbookRender(workbook, options);
        renderer.ToImage("output_white_background.tiff");

        Console.WriteLine("TIFF image generated with a white background.");
    }
}
