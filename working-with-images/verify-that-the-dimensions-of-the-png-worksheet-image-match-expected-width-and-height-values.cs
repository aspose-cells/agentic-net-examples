// Title: Check exported worksheet PNG size matches expected dimensions using Aspose.Cells in C#
// Common Searches: C# read PNG header to obtain image width and height without System.Drawing | Aspose.Cells SheetRender verify output PNG dimensions programmatically | how to compare expected worksheet image size after exporting to PNG in .NET | determine PNG dimensions of Excel sheet rendered with Aspose.Cells | validate PNG resolution and size after setting ImageOrPrintOptions DPI
// Tags: Aspose.Cells worksheet to PNG rendering | extract PNG width height from header C# | validate exported image dimensions | ImageOrPrintOptions resolution configuration | compare actual PNG size with expected values

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, renders the first worksheet to a PNG file at 96 DPI using Aspose.Cells SheetRender, then reads the PNG header to retrieve its width and height and compares them against the expected 800 × 600 pixels, reporting whether the dimensions match.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string pngPath = "worksheet.png";

        try
        {
            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Set image rendering options (96 DPI). PNG is the default format.
            var imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 96,
                VerticalResolution = 96
            };

            // Render the first worksheet to a PNG image
            var sheetRender = new SheetRender(workbook.Worksheets[0], imgOptions);
            sheetRender.ToImage(0, pngPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during workbook processing: {ex.Message}");
            return;
        }

        try
        {
            // Verify the dimensions of the generated PNG without using System.Drawing
            if (!File.Exists(pngPath))
            {
                Console.WriteLine($"Generated image not found: {pngPath}");
                return;
            }

            if (TryGetPngDimensions(pngPath, out int actualWidth, out int actualHeight))
            {
                const int expectedWidth = 800;   // expected width
                const int expectedHeight = 600;  // expected height

                if (actualWidth == expectedWidth && actualHeight == expectedHeight)
                {
                    Console.WriteLine("PNG dimensions match expected values.");
                }
                else
                {
                    Console.WriteLine($"Dimension mismatch. Expected {expectedWidth}x{expectedHeight}, got {actualWidth}x{actualHeight}.");
                }
            }
            else
            {
                Console.WriteLine("Failed to read PNG dimensions.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error while verifying image: {ex.Message}");
        }
    }

    // Reads PNG header to obtain width and height (big‑endian integers)
    private static bool TryGetPngDimensions(string filePath, out int width, out int height)
    {
        width = height = 0;
        try
        {
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            using (var br = new BinaryReader(fs))
            {
                // PNG signature (8 bytes)
                var signature = br.ReadBytes(8);
                if (signature.Length != 8 ||
                    signature[0] != 0x89 || signature[1] != 0x50 || signature[2] != 0x4E ||
                    signature[3] != 0x47 || signature[4] != 0x0D || signature[5] != 0x0A ||
                    signature[6] != 0x1A || signature[7] != 0x0A)
                {
                    return false; // Not a PNG file
                }

                // Read the IHDR chunk length and type
                int length = ReadInt32BigEndian(br);
                var chunkType = new string(br.ReadChars(4));
                if (chunkType != "IHDR")
                {
                    return false; // Unexpected chunk
                }

                // Width and height are the first two 4‑byte integers in IHDR
                width = ReadInt32BigEndian(br);
                height = ReadInt32BigEndian(br);
                return true;
            }
        }
        catch
        {
            return false;
        }
    }

    // Helper to read a 4‑byte big‑endian integer
    private static int ReadInt32BigEndian(BinaryReader br)
    {
        var bytes = br.ReadBytes(4);
        if (bytes.Length < 4) throw new EndOfStreamException();
        return (bytes[0] << 24) | (bytes[1] << 16) | (bytes[2] << 8) | bytes[3];
    }
}
