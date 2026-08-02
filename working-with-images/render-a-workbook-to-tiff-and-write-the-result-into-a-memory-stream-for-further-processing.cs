using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing; // for ImageType enum

class RenderWorkbookToTiffStream
{
    public static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet ws = workbook.Worksheets[0];
        ws.Cells["A1"].PutValue("Aspose.Cells TIFF Rendering Demo");
        ws.Cells["A2"].PutValue(DateTime.Now);

        // Configure image rendering options for TIFF output
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Tiff;                     // Specify TIFF format
        options.TiffCompression = TiffCompression.CompressionLZW; // Optional compression
        options.OnePagePerSheet = true;                         // Render each sheet as a single page

        // Create a renderer for the whole workbook
        WorkbookRender renderer = new WorkbookRender(workbook, options);

        // Render the workbook to a memory stream
        using (MemoryStream tiffStream = new MemoryStream())
        {
            renderer.ToImage(tiffStream); // Render entire workbook as TIFF into the stream

            // Reset stream position if further processing is needed
            tiffStream.Position = 0;

            // Example: save the stream to a file (optional, for verification)
            using (FileStream file = new FileStream("output.tiff", FileMode.Create, FileAccess.Write))
            {
                tiffStream.CopyTo(file);
            }

            Console.WriteLine($"TIFF image rendered to stream. Stream length: {tiffStream.Length} bytes");
        }

        // Release resources used by the renderer
        renderer.Dispose();
    }
}