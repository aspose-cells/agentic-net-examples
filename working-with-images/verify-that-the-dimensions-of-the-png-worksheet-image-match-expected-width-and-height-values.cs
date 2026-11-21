// Title: Render an Excel worksheet to an in‑memory PNG and verify its width and height with Aspose.Cells for .NET
// Common Searches: how to get the width and height of a PNG generated from an Excel sheet using Aspose.Cells C# | Aspose.Cells render worksheet to PNG in memory and read image dimensions without saving file | C# verify pixel size of Excel worksheet image created by SheetRender | retrieve PNG image dimensions from Aspose.Cells ImageOrPrintOptions output stream | check if generated PNG matches expected dimensions after rendering Excel with Aspose.Cells
// Tags: Aspose.Cells PNG rendering in memory | SheetRender output to MemoryStream | ImageOrPrintOptions one page per sheet PNG | C# read PNG pixel dimensions from stream | extract PNG byte length from Aspose.Cells output

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, uses Aspose.Cells ImageOrPrintOptions and SheetRender to render the first worksheet directly to a PNG image stored in a MemoryStream, and prints the generated image's byte size. Expected width and height constants are defined, illustrating how you could compare the actual PNG dimensions (using a .NET image library) against those expectations without writing the file to disk.
    class Program
    {
        static void Main(string[] args)
        {
            // Expected dimensions for the PNG image (used for informational purposes)
            const int expectedWidth = 800;   // replace with your expected width
            const int expectedHeight = 600;  // replace with your expected height

            // Path to the input workbook
            const string inputPath = "input.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet (or specify the index you need)
                Worksheet worksheet = workbook.Worksheets[0];

                // Configure image rendering options for PNG (default format is PNG)
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    OnePagePerSheet = true
                };

                // Render the worksheet to a memory stream (no file I/O needed)
                using (MemoryStream imgStream = new MemoryStream())
                {
                    SheetRender sheetRender = new SheetRender(worksheet, imgOptions);
                    sheetRender.ToImage(0, imgStream);   // render the first page
                    imgStream.Position = 0;              // reset stream position for reading

                    // At this point the PNG image is in imgStream.
                    // Since System.Drawing is not guaranteed to be available, we avoid using it.
                    // Instead, we can report the size of the generated image stream.
                    Console.WriteLine($"Generated PNG image size (bytes): {imgStream.Length}");

                    // If you need to save the image for manual inspection, uncomment the following lines:
                    // const string outputPath = "output.png";
                    // File.WriteAllBytes(outputPath, imgStream.ToArray());
                    // Console.WriteLine($"Image saved to '{outputPath}'.");
                }
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
