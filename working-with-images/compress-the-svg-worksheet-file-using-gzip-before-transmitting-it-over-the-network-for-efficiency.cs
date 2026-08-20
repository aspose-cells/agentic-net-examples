// Title: Compress Aspose.Cells SVG Worksheet with GZip in C#
// Description: Shows how to render a worksheet to SVG using Aspose.Cells and then compress the SVG file with GZipStream to lower size and speed up network transmission.
// Keywords: Aspose.Cells | C# | SVG rendering | GZip compression | GZipStream | worksheet to SVG | compress SVG | network transfer | file compression .NET | Aspose.Cells SVG output
// Common Searches: Aspose.Cells render worksheet to SVG C# | How to gzip an SVG file in .NET | Compress Aspose.Cells SVG output | C# GZipStream example for SVG | Reduce size of Excel SVG export
// Developer Intent: Generate an SVG representation of an Excel worksheet with Aspose.Cells and then compress that SVG using GZip so it can be sent over the network with minimal bandwidth.
// Use Cases: Web preview where a compressed SVG snapshot of a spreadsheet is delivered to browsers. | REST API that returns a GZip‑compressed SVG of a workbook for client‑side rendering. | Archiving multiple worksheet SVGs in a .gz package for long‑term storage. | Streaming compressed SVG to mobile apps to conserve data usage.
// AI Prompts: Write a C# method that accepts a Workbook, renders the first worksheet to SVG in memory, and returns a GZip‑compressed byte array. | Provide C# code that renders an Aspose.Cells worksheet to SVG, compresses it with GZipStream, and safely deletes temporary files. | Explain the performance advantages of gzipping SVG output from Aspose.Cells and how to set appropriate HTTP response headers for compressed content.

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to render a worksheet to SVG using Aspose.Cells and then compress the SVG file with GZipStream to lower size and speed up network transmission.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Fruit");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(150);
            worksheet.Cells["A3"].PutValue("Orange");
            worksheet.Cells["B3"].PutValue(250);

            // Configure SVG rendering options (FitToViewPort ensures the whole sheet fits the viewport)
            SvgImageOptions svgOptions = new SvgImageOptions
            {
                FitToViewPort = true
            };

            // Render the worksheet to an SVG file
            string svgFilePath = "worksheet.svg";
            SheetRender renderer = new SheetRender(worksheet, svgOptions);
            renderer.ToImage(0, svgFilePath);

            // Verify that the SVG file was created before compression
            if (!File.Exists(svgFilePath))
                throw new FileNotFoundException("SVG file was not created.", svgFilePath);

            // Compress the generated SVG using GZip
            string gzFilePath = "worksheet.svg.gz";
            using (FileStream originalFile = File.OpenRead(svgFilePath))
            using (FileStream compressedFile = File.Create(gzFilePath))
            using (GZipStream gzipStream = new GZipStream(compressedFile, CompressionMode.Compress))
            {
                originalFile.CopyTo(gzipStream);
            }

            Console.WriteLine($"SVG file has been compressed to: {gzFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
