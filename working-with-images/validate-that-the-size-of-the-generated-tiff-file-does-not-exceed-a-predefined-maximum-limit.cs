// Title: Check that a TIFF image generated from an Excel workbook does not exceed a 5 MB limit using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a workbook, saves it as a TIFF with Aspose.Cells, and throws an exception if the saved file size exceeds a configurable byte limit. | Add logic to automatically delete the TIFF file when its size is larger than the defined maximum after the Aspose.Cells export. | Modify the example to log the actual file size and the allowed limit, returning a boolean that indicates whether the TIFF meets the size requirement.
// Common Searches: Aspose.Cells .NET how to enforce maximum TIFF file size after export | C# verify size of TIFF saved from Excel workbook using Aspose.Cells | prevent oversized TIFF image when converting Excel to image with Aspose.Cells | check file size of generated TIFF in C# Aspose.Cells example
// Tags: Aspose.Cells save workbook as TIFF size limit | C# validate TIFF file size after export | maximum TIFF file size check Aspose.Cells | file size verification for generated TIFF .NET | configure TIFF size threshold Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program creates a simple workbook, exports it to a TIFF image with Aspose.Cells, then checks the resulting file size against a 5 MB threshold, reporting success or an error and optionally handling oversized files.
class Program
{
    static void Main()
    {
        try
        {
            // Maximum allowed TIFF file size (5 MB)
            const long MaxTiffSizeBytes = 5 * 1024 * 1024;

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["A3"].PutValue(456);

            // Path for the generated TIFF file
            string tiffPath = "output.tiff";

            // Save the workbook as a TIFF image (one page per sheet by default)
            workbook.Save(tiffPath, SaveFormat.Tiff);

            // Verify that the file was created
            if (!File.Exists(tiffPath))
            {
                Console.WriteLine("Error: TIFF file was not created.");
                return;
            }

            // Check the file size
            FileInfo fileInfo = new FileInfo(tiffPath);
            if (fileInfo.Length > MaxTiffSizeBytes)
            {
                Console.WriteLine($"Error: TIFF size {fileInfo.Length} bytes exceeds the limit of {MaxTiffSizeBytes} bytes.");
                // Optional: delete the oversized file
                // File.Delete(tiffPath);
            }
            else
            {
                Console.WriteLine($"TIFF saved successfully. Size: {fileInfo.Length} bytes.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
