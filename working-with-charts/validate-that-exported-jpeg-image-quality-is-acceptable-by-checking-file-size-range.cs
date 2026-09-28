// Title: Check exported worksheet image size falls within 10KB‑500KB using Aspose.Cells for .NET
// AI Prompts: Write C# code that uses Aspose.Cells to render a worksheet to a PNG image, saves it, and confirms the file size is between 10 KB and 500 KB. | Show how to configure ImageOrPrintOptions, render the first sheet page, and programmatically validate the exported image's byte size with Aspose.Cells in a .NET application.
// Common Searches: C# Aspose.Cells export worksheet to PNG and ensure file size is within a specific range | How to validate image quality of an exported Excel sheet by checking its byte size in .NET | Aspose.Cells render sheet to image and verify size limits 10KB 500KB | Check exported chart image size using Aspose.Cells C# example | Validate PNG export size after rendering Excel worksheet with Aspose.Cells
// Tags: Aspose.Cells worksheet PNG export size validation | C# ImageOrPrintOptions file size verification | exported Excel sheet image byte range check | Aspose.Cells render sheet to image size constraint | validate image quality by file size .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// // This program creates a workbook, renders the first worksheet page to a PNG image using Aspose.Cells, writes the image to disk, and then validates that the image's byte size falls between 10 KB and 500 KB, reporting whether the size is acceptable.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Header");
            sheet.Cells["A2"].PutValue(100);
            sheet.Cells["B2"].PutValue(200);
            sheet.Cells["C2"].PutValue(300);

            // Configure image rendering options (default PNG format)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

            // Render the worksheet to an image
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            using (MemoryStream ms = new MemoryStream())
            {
                // Render the first page (index 0) to the memory stream
                renderer.ToImage(0, ms);
                ms.Position = 0; // Reset stream position

                // Define output file path (PNG format)
                string outputFile = "ExportedSheet.png";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Write image to disk with error handling
                try
                {
                    File.WriteAllBytes(outputFile, ms.ToArray());
                }
                catch (Exception writeEx)
                {
                    Console.WriteLine($"Failed to write image file: {writeEx.Message}");
                }

                // Validate image file size (10KB – 500KB)
                long fileSize = ms.Length;
                const long minSize = 10 * 1024;   // 10 KB
                const long maxSize = 500 * 1024; // 500 KB

                if (fileSize < minSize || fileSize > maxSize)
                {
                    Console.WriteLine($"Image size {fileSize} bytes is outside the acceptable range.");
                }
                else
                {
                    Console.WriteLine($"Image size {fileSize} bytes is within the acceptable range.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
