// Title: Validate Exported Chart PNG Size and Format with Aspose.Cells for .NET
// Description: Creates a workbook, adds a column chart, exports it to PNG using ImageOrPrintOptions with a 800 × 600 size, reads the PNG header to obtain actual width, height, and signature, and reports whether the image matches the expected dimensions and format.
// Keywords: Aspose.Cells | C# | export chart PNG | validate PNG dimensions | ImageOrPrintOptions SetDesiredSize | PNG header parsing | chart image verification | Aspose.Cells chart export
// Common Searches: Aspose.Cells verify exported PNG size | C# read PNG width and height from file | Set exact PNG dimensions when exporting chart Aspose | Check if exported file is a valid PNG Aspose.Cells | Automated test for chart image size C#
// Developer Intent: Ensure that a chart exported via Aspose.Cells is a valid PNG file with the exact width and height specified in the export options.
// Use Cases: Automated CI test that confirms chart images meet size requirements before deployment. | Pre‑publish validation to guarantee exported PNGs match layout specifications for reports. | Integration step that checks PNG validity before feeding images into downstream processing pipelines.
// AI Prompts: Write a C# method that compares expected and actual PNG dimensions after exporting a chart with Aspose.Cells and returns a detailed validation result. | Generate code that logs a warning and aborts the workflow if the exported PNG size does not match the desired dimensions. | Suggest improvements to the PNG validation logic, such as supporting multiple image formats or leveraging Aspose.Imaging for more robust verification.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Creates a workbook, adds a column chart, exports it to PNG using ImageOrPrintOptions with a 800 × 600 size, reads the PNG header to obtain actual width, height, and signature, and reports whether the image matches the expected dimensions and format.
class ValidateExportedPng
{
    static void Main()
    {
        // Expected dimensions
        const int expectedWidth = 800;
        const int expectedHeight = 600;

        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIdx];
            chart.SetChartDataRange("A1:B4", true);

            // Set image options: PNG format and desired size
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
            // The default format is PNG when the file extension is .png, so explicit setting is optional.
            imgOptions.SetDesiredSize(expectedWidth, expectedHeight, false); // keepAspectRatio = false

            // Export chart to PNG file using the options
            string pngPath = "exported_chart.png";

            try
            {
                chart.ToImage(pngPath, imgOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error exporting chart to PNG: {ex.Message}");
                return;
            }

            // Validate the exported PNG file
            if (!File.Exists(pngPath))
            {
                Console.WriteLine($"File not found: {pngPath}");
                return;
            }

            try
            {
                (int actualWidth, int actualHeight) = GetPngDimensions(pngPath);
                bool sizeMatches = actualWidth == expectedWidth && actualHeight == expectedHeight;
                bool formatMatches = IsPngFormat(pngPath);

                Console.WriteLine($"Image path: {pngPath}");
                Console.WriteLine($"Expected size: {expectedWidth}x{expectedHeight}");
                Console.WriteLine($"Actual size: {actualWidth}x{actualHeight}");
                Console.WriteLine($"Size match: {sizeMatches}");
                Console.WriteLine($"Expected format: PNG");
                Console.WriteLine($"Actual format: {(formatMatches ? "PNG" : "Unknown")}");
                Console.WriteLine($"Format match: {formatMatches}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error validating PNG file: {ex.Message}");
            }

            // Optional: clean up the generated file
            // File.Delete(pngPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }

    // Reads PNG header to extract width and height (big‑endian)
    private static (int width, int height) GetPngDimensions(string filePath)
    {
        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        using (BinaryReader br = new BinaryReader(fs))
        {
            // PNG signature (8 bytes)
            byte[] signature = br.ReadBytes(8);
            if (signature.Length != 8 ||
                signature[0] != 137 || signature[1] != 80 || signature[2] != 78 ||
                signature[3] != 71 || signature[4] != 13 || signature[5] != 10 ||
                signature[6] != 26 || signature[7] != 10)
            {
                throw new InvalidDataException("File is not a valid PNG.");
            }

            // Read IHDR chunk length and type
            uint ihdrLength = ReadBigEndianUInt32(br);
            string ihdrType = new string(br.ReadChars(4));
            if (ihdrType != "IHDR")
                throw new InvalidDataException("IHDR chunk not found.");

            // Width and height are the next 8 bytes
            uint width = ReadBigEndianUInt32(br);
            uint height = ReadBigEndianUInt32(br);

            return ((int)width, (int)height);
        }
    }

    // Checks PNG signature
    private static bool IsPngFormat(string filePath)
    {
        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            byte[] signature = new byte[8];
            int read = fs.Read(signature, 0, 8);
            return read == 8 &&
                   signature[0] == 137 && signature[1] == 80 && signature[2] == 78 &&
                   signature[3] == 71 && signature[4] == 13 && signature[5] == 10 &&
                   signature[6] == 26 && signature[7] == 10;
        }
    }

    // Helper to read a 4‑byte unsigned int in big‑endian order
    private static uint ReadBigEndianUInt32(BinaryReader br)
    {
        byte[] bytes = br.ReadBytes(4);
        if (bytes.Length != 4)
            throw new EndOfStreamException();
        return ((uint)bytes[0] << 24) |
               ((uint)bytes[1] << 16) |
               ((uint)bytes[2] << 8) |
               bytes[3];
    }
}
