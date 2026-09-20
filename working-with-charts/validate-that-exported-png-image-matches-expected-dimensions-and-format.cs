// Title: Validate PNG image dimensions after exporting the first worksheet with Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to render the first worksheet of an Excel file to a PNG file at 96 dpi, then reads the PNG header to confirm its width and height match expected values. | Create a C# helper that parses the IHDR chunk of a PNG file to verify the format and extract dimensions, and integrate it with Aspose.Cells sheet rendering for automated validation.
// Common Searches: how to check size of PNG exported from Excel using Aspose.Cells C# | C# verify dimensions of worksheet image rendered by Aspose.Cells | read PNG IHDR chunk to get width and height in .NET | set DPI for Excel to PNG conversion with Aspose.Cells
// Tags: Aspose.Cells worksheet to PNG export | PNG dimension validation in C# | read PNG IHDR chunk .NET | set image DPI Aspose.Cells rendering | automated image size check after Excel export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, renders the first worksheet to a PNG file at 96 dpi using Aspose.Cells, then opens the PNG, validates its signature, reads the IHDR chunk to obtain width and height, and confirms the image matches the expected 800 × 600 pixels.
class Program
{
    static void Main()
    {
        try
        {
            // Expected image dimensions
            const int expectedWidth = 800;
            const int expectedHeight = 600;

            // Input workbook path
            string inputPath = "input.xlsx";
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The workbook does not contain any worksheets.");
                return;
            }

            // Configure image export options (default format is PNG)
            var options = new ImageOrPrintOptions
            {
                HorizontalResolution = 96,
                VerticalResolution = 96
            };

            // Render the first worksheet to an image file
            var sheet = workbook.Worksheets[0];
            var renderer = new SheetRender(sheet, options);
            string outputPath = "exported.png";

            try
            {
                renderer.ToImage(0, outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during rendering: {ex.Message}");
                return;
            }

            // Validate the exported image
            if (!File.Exists(outputPath))
            {
                Console.WriteLine($"Exported image not found: {outputPath}");
                return;
            }

            if (ValidatePng(outputPath, expectedWidth, expectedHeight))
            {
                Console.WriteLine("Validation succeeded: PNG image has the expected dimensions and format.");
            }
            else
            {
                Console.WriteLine("Validation failed: PNG image does not match expected dimensions or format.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }

    // Validates that a file is a PNG and checks its width and height from the IHDR chunk.
    private static bool ValidatePng(string filePath, int expectedWidth, int expectedHeight)
    {
        try
        {
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            using (var br = new BinaryReader(fs))
            {
                // PNG signature (8 bytes)
                byte[] signature = br.ReadBytes(8);
                byte[] pngSignature = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
                for (int i = 0; i < 8; i++)
                {
                    if (signature[i] != pngSignature[i])
                        return false; // Not a PNG file
                }

                // Read the first chunk header (length + type)
                uint length = ReadBigEndianUInt32(br);
                string chunkType = new string(br.ReadChars(4));

                if (chunkType != "IHDR")
                    return false; // Unexpected first chunk

                // IHDR data: width (4 bytes), height (4 bytes), etc.
                uint width = ReadBigEndianUInt32(br);
                uint height = ReadBigEndianUInt32(br);

                // Compare dimensions
                return width == (uint)expectedWidth && height == (uint)expectedHeight;
            }
        }
        catch
        {
            // Any error during parsing means validation failed
            return false;
        }
    }

    // Reads a 4‑byte unsigned integer in big‑endian order.
    private static uint ReadBigEndianUInt32(BinaryReader br)
    {
        byte[] bytes = br.ReadBytes(4);
        if (bytes.Length < 4)
            throw new EndOfStreamException();
        return ((uint)bytes[0] << 24) |
               ((uint)bytes[1] << 16) |
               ((uint)bytes[2] << 8) |
               bytes[3];
    }
}
