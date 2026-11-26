// Title: Create a multi‑page TIFF from the first three worksheets of an Excel file using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export first three Excel sheets to a multi‑page TIFF | How to copy selected worksheets to a new workbook before rendering to TIFF in .NET | Aspose.Cells render workbook to TIFF with LZW compression and 300 DPI | Generate multi‑page TIFF from specific worksheets using Aspose.Cells ImageOrPrintOptions | Save Excel workbook as multi‑page TIFF with selected sheets in C#
// Tags: selective worksheet copy Aspose.Cells | TIFF LZW compression settings ImageOrPrintOptions | multi‑page TIFF rendering with WorkbookRender | render specific Excel sheets to TIFF .NET | export Excel worksheets as multi‑page TIFF C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;          // For ImageType enum
using Aspose.Cells.Rendering;       // For rendering classes

// The example loads an Excel workbook, copies the first three worksheets into a new workbook, configures ImageOrPrintOptions for TIFF with LZW compression and 300 dpi, and uses WorkbookRender to generate a multi‑page TIFF file.
class Program
{
    static void Main(string[] args)
    {
        // Paths – adjust as needed
        string inputFile = "input.xlsx";
        string outputFile = "output.tiff";

        // Validate input file
        if (!File.Exists(inputFile))
        {
            Console.WriteLine($"Error: Input file \"{inputFile}\" does not exist.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook sourceWb = new Workbook(inputFile);

            // Prepare a new workbook that will contain only the first three worksheets
            Workbook targetWb = new Workbook();
            // Remove the default empty sheet created by the constructor
            if (targetWb.Worksheets.Count > 0)
                targetWb.Worksheets.RemoveAt(0);

            int sheetsToCopy = Math.Min(3, sourceWb.Worksheets.Count);
            for (int i = 0; i < sheetsToCopy; i++)
            {
                // Copy the worksheet by name (preserves data, formatting, page setup, etc.)
                targetWb.Worksheets.AddCopy(sourceWb.Worksheets[i].Name);
            }

            // Configure image rendering options for TIFF
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Tiff,                         // Output format
                TiffCompression = TiffCompression.CompressionLZW, // LZW compression
                TiffColorDepth = ColorDepth.Format24bpp,            // 24‑bit color depth
                HorizontalResolution = 300,                         // DPI
                VerticalResolution = 300
            };

            // Render the workbook to a multi‑page TIFF file
            WorkbookRender renderer = new WorkbookRender(targetWb, imgOptions);
            renderer.ToImage(outputFile); // Renders whole workbook as a multi‑page TIFF

            Console.WriteLine($"Multi‑page TIFF generated successfully: {outputFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while generating the TIFF:");
            Console.WriteLine(ex.Message);
        }
    }
}
