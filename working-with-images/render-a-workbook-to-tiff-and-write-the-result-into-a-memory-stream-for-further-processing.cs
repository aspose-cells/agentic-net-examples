// Title: Render an Excel Workbook to TIFF in a MemoryStream with Aspose.Cells (C#)
// Description: This example creates a Workbook, sets TIFF options (LZW compression, one page per sheet), and uses WorkbookRender.ToImage(Stream) to write the image directly to a MemoryStream. The stream can then be reset, saved, sent over a network, or stored without creating a temporary file.
// Keywords: Aspose.Cells TIFF rendering | C# MemoryStream image export | WorkbookRender ToImage stream | LZW compression TIFF | Excel to TIFF in memory | export Excel as image C# | Aspose.Cells image options
// Common Searches: Aspose.Cells render workbook to TIFF memory stream C# | How to export Excel as TIFF without a file using Aspose.Cells | C# convert worksheet to TIFF image stream | Aspose.Cells LZW compression TIFF example | Create multi‑page TIFF from Excel in memory
// Developer Intent: The developer needs to convert an Excel workbook to a TIFF image and keep the result in a MemoryStream for further processing such as sending over a web API, storing in a database, or attaching to an email.
// Use Cases: Generate a TIFF preview of a report and embed it in an email attachment without writing to disk. | Send the TIFF byte stream to a third‑party printing service that accepts image streams. | Store the TIFF bytes in a database column for archival of generated Excel reports.
// AI Prompts: Show how to render each worksheet as a separate page in a multi‑page TIFF stored in a MemoryStream. | Provide code for reading the TIFF bytes from the MemoryStream and returning them from an ASP.NET Core controller action. | Explain how to switch the compression to CCITT Group 4 and retrieve the resulting byte array.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsTiffMemoryStreamDemo
{
    // This example creates a Workbook, sets TIFF options (LZW compression, one page per sheet), and uses WorkbookRender.ToImage(Stream) to write the image directly to a MemoryStream. The stream can then be reset, saved, sent over a network, or stored without creating a temporary file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells TIFF rendering to MemoryStream");
            sheet.Cells["A2"].PutValue(DateTime.Now);

            // Configure rendering options for TIFF output
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Tiff,                 // Specify TIFF format
                TiffCompression = TiffCompression.CompressionLZW, // Optional: set compression
                OnePagePerSheet = true                      // Render each sheet as a single page
            };

            // Create a renderer for the whole workbook
            WorkbookRender renderer = new WorkbookRender(workbook, renderOptions);

            // Render the workbook to a memory stream
            using (MemoryStream tiffStream = new MemoryStream())
            {
                renderer.ToImage(tiffStream); // Uses WorkbookRender.ToImage(Stream)

                // The stream now contains the TIFF image data.
                // Reset position if the stream will be read later.
                tiffStream.Position = 0;

                // Example: write the stream to a file (optional, for verification)
                using (FileStream file = new FileStream("output.tiff", FileMode.Create, FileAccess.Write))
                {
                    tiffStream.CopyTo(file);
                }

                Console.WriteLine($"TIFF image rendered to memory stream. Length = {tiffStream.Length} bytes.");
            }

            // Clean up
            renderer.Dispose();
        }
    }
}
