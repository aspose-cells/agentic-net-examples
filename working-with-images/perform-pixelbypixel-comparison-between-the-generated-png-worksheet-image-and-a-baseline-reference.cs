// Title: Generate a PNG image of a worksheet with Aspose.Cells in C# and compare it pixel‑by‑pixel to a baseline file
// AI Prompts: Render the first page of a workbook to a PNG file using Aspose.Cells SheetRender, then load the PNG and a reference PNG and count differing bytes to confirm visual equality. | Create a C# method that takes two image paths, checks for length mismatches, iterates through their byte arrays, and returns a boolean plus the number of mismatched bytes indicating whether the images are identical.
// Common Searches: how to render an Excel worksheet to PNG with Aspose.Cells and verify the output in C# | C# compare generated worksheet image to baseline PNG pixel by pixel | Aspose.Cells SheetRender compare rendered image with reference file | byte level comparison of two PNG files produced from Excel in .NET | detect visual differences between Excel sheet screenshots using Aspose.Cells
// Tags: worksheet to PNG rendering Aspose.Cells | byte‑by‑byte PNG comparison C# | image equality verification Aspose.Cells | pixel level diff of Excel sheet image | SheetRender PNG output validation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook, adds sample data, and uses Aspose.Cells SheetRender with ImageOrPrintOptions to save the first worksheet page as a PNG file. It then ensures both the generated PNG and a baseline PNG exist, reads their byte arrays, and performs a byte‑by‑byte comparison (treating length differences as a mismatch) to determine if the images are identical, reporting the result and the count of mismatched bytes.
class WorksheetImageComparer
{
    static void Main()
    {
        try
        {
            // Create a workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "SampleSheet";
            sheet.Cells["A1"].PutValue("Aspose");
            sheet.Cells["B2"].PutValue(12345);
            sheet.Cells["C3"].PutValue(DateTime.Now);

            // Path for the generated image
            string generatedImagePath = "generated.png";

            // Configure rendering options (resolution only; format inferred from file extension)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 96,
                VerticalResolution = 96
            };

            // Render the first page of the worksheet to an image file
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            renderer.ToImage(0, generatedImagePath); // 0 = first page

            // Verify that the generated image was created
            if (!File.Exists(generatedImagePath))
            {
                Console.WriteLine($"Failed to generate image at '{generatedImagePath}'.");
                return;
            }

            // Path to the baseline reference image
            string baselineImagePath = "baseline.png";

            // Ensure baseline image exists
            if (!File.Exists(baselineImagePath))
            {
                Console.WriteLine($"Baseline image not found at '{baselineImagePath}'.");
                return;
            }

            // Perform byte‑by‑byte comparison (approximates pixel comparison)
            bool imagesAreEqual = CompareImagesByteByByte(generatedImagePath, baselineImagePath, out int mismatchedBytes);

            // Output the result
            Console.WriteLine($"Images are equal: {imagesAreEqual}");
            Console.WriteLine($"Mismatched byte count: {mismatchedBytes}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    /// <param name="pathA">Path to the first image.</param>
    /// <param name="pathB">Path to the second image.</param>
    /// <param name="mismatchedBytes">Number of bytes that differ.</param>
    /// <returns>True if images are identical; otherwise false.</returns>
    static bool CompareImagesByteByByte(string pathA, string pathB, out int mismatchedBytes)
    {
        mismatchedBytes = 0;

        try
        {
            byte[] bytesA = File.ReadAllBytes(pathA);
            byte[] bytesB = File.ReadAllBytes(pathB);

            // Different lengths mean images differ
            if (bytesA.Length != bytesB.Length)
            {
                mismatchedBytes = -1; // indicate length mismatch
                return false;
            }

            for (int i = 0; i < bytesA.Length; i++)
            {
                if (bytesA[i] != bytesB[i])
                {
                    mismatchedBytes++;
                }
            }

            return mismatchedBytes == 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during image comparison: {ex.Message}");
            mismatchedBytes = -1;
            return false;
        }
    }
}
