// Title: Compress an Aspose.Cells worksheet saved as SVG with GZip in C# for faster network transfer
// AI Prompts: Generate C# code that uses Aspose.Cells to export a worksheet to SVG, then compress the SVG bytes with GZipStream and save as a .gz file. | Demonstrate how to pipe the SVG MemoryStream from Aspose.Cells directly into a GZipStream for network‑ready compression in C#. | Show how to configure ImageSaveOptions for SVG and perform in‑memory GZip compression before writing the result to disk.
// Common Searches: C# how to gzip compress SVG output from Aspose.Cells before sending over HTTP | Aspose.Cells save worksheet as SVG and compress with GZipStream for web API | example of streaming Aspose.Cells SVG to GZipStream in .NET
// Tags: Aspose.Cells export worksheet to SVG | C# GZipStream compress SVG data | in‑memory SVG compression with Aspose.Cells | network‑optimized SVG file from Excel workbook

using Aspose.Cells;
using Aspose.Cells.Saving;
using System;
using System.IO;
using System.IO.Compression;

// The sample creates a workbook, fills it with data, saves the first worksheet as an SVG image using ImageSaveOptions, compresses the SVG stream with GZipStream in memory, and writes the compressed .gz file, enabling efficient transmission over the network.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and access the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);

            // Set SVG save options using the recommended ImageSaveOptions
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Svg);
            saveOptions.ImageOrPrintOptions.OnePagePerSheet = true;

            // Save the worksheet to a memory stream as SVG
            using (MemoryStream svgStream = new MemoryStream())
            {
                workbook.Save(svgStream, saveOptions);
                svgStream.Position = 0; // Reset for reading

                // Compress the SVG data using GZip
                byte[] compressedData;
                using (MemoryStream compressedStream = new MemoryStream())
                {
                    using (GZipStream gzip = new GZipStream(compressedStream, CompressionMode.Compress, true))
                    {
                        svgStream.CopyTo(gzip);
                    }
                    compressedData = compressedStream.ToArray();
                }

                // Write the compressed SVG to a file
                File.WriteAllBytes("Worksheet.svg.gz", compressedData);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
