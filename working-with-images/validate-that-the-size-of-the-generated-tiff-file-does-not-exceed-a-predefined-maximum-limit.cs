// Title: Validate TIFF Output Size with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills it with sample data, renders the first worksheet to a TIFF stream using LZW compression at 150 dpi, checks the stream length against a 5 MB limit, throws an exception if the limit is exceeded, and saves the file when the size is acceptable.
// Keywords: Aspose.Cells TIFF size check | C# validate TIFF file size | limit TIFF output Aspose | render worksheet to TIFF .NET | TIFF compression LZW Aspose.Cells
// Common Searches: Aspose.Cells how to limit TIFF file size in C# | check TIFF size after rendering with Aspose.Cells | C# code to validate generated TIFF does not exceed 5 MB | prevent oversized TIFF export using Aspose.Cells
// Developer Intent: Ensure the TIFF image generated from a worksheet stays within a predefined size constraint before persisting it.
// Use Cases: Avoid email attachment rejections caused by large TIFF exports. | Enforce storage quotas in document management systems by rejecting oversized TIFFs. | Optimize export settings (resolution, compression) to meet size limits for regulatory compliance.
// AI Prompts: Write C# code with Aspose.Cells that renders a worksheet to a TIFF stream and raises an error if the stream exceeds 5 MB. | Suggest ImageOrPrintOptions adjustments (resolution, compression) to reduce TIFF size below a target threshold. | Create a unit test that verifies the TIFF size validation logic for a workbook with many rows.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTiffSizeValidation
{
    // Creates a workbook, fills it with sample data, renders the first worksheet to a TIFF stream using LZW compression at 150 dpi, checks the stream length against a 5 MB limit, throws an exception if the limit is exceeded, and saves the file when the size is acceptable.
    class Program
    {
        // Maximum allowed TIFF file size (5 MB)
        private const long MaxTiffSizeBytes = 5 * 1024 * 1024; // 5,242,880 bytes

        static void Main()
        {
            try
            {
                // Create a new workbook and obtain the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate worksheet with sample data
                worksheet.Cells["A1"].PutValue("Aspose.Cells TIFF Size Validation Demo");
                for (int row = 2; row <= 100; row++)
                {
                    worksheet.Cells[row - 1, 0].PutValue($"Row {row}");
                    worksheet.Cells[row - 1, 1].PutValue(row * 10);
                }

                // Configure rendering options for TIFF output
                ImageOrPrintOptions options = new ImageOrPrintOptions
                {
                    // ImageFormat property is not required for ToTiff; remove invalid assignment
                    TiffCompression = TiffCompression.CompressionLZW, // LZW compression
                    HorizontalResolution = 150,
                    VerticalResolution = 150,
                    OnePagePerSheet = true // Single page per sheet
                };

                // Render the worksheet to a TIFF image in memory
                SheetRender renderer = new SheetRender(worksheet, options);
                using (MemoryStream tiffStream = new MemoryStream())
                {
                    try
                    {
                        renderer.ToTiff(tiffStream);
                    }
                    catch (Exception renderEx)
                    {
                        Console.WriteLine($"Rendering failed: {renderEx.Message}");
                        return;
                    }

                    long tiffSize = tiffStream.Length;
                    Console.WriteLine($"Generated TIFF size: {tiffSize} bytes");

                    if (tiffSize > MaxTiffSizeBytes)
                    {
                        Console.WriteLine($"Error: TIFF size exceeds the maximum allowed limit of {MaxTiffSizeBytes} bytes.");
                        throw new InvalidOperationException("Generated TIFF file is too large.");
                    }

                    // Save TIFF to disk if size is acceptable
                    string outputPath = "output_valid.tiff";

                    // Ensure target directory exists
                    string directory = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    try
                    {
                        File.WriteAllBytes(outputPath, tiffStream.ToArray());
                        Console.WriteLine($"TIFF file saved successfully to '{outputPath}'.");
                    }
                    catch (Exception ioEx)
                    {
                        Console.WriteLine($"Failed to write TIFF file: {ioEx.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
