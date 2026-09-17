// Title: How to render an Excel worksheet to PNG with Aspose.Cells and perform a pixel‑by‑pixel comparison against a baseline image in C#
// AI Prompts: Generate C# code that uses Aspose.Cells to export a worksheet as a PNG file and then compares the PNG to a reference image pixel by pixel, returning the count of differing pixels. | Create a method that reads two PNG files into byte arrays, iterates through the bytes, counts mismatches, and reports whether the images are identical. | Add robust error handling for missing files, size mismatches, and I/O exceptions when validating rendered worksheet images in automated tests.
// Common Searches: Aspose.Cells C# render worksheet to PNG and compare with expected image | pixel level regression test for Excel sheet image output using Aspose.Cells | C# byte array comparison of two PNG files generated from Excel | how to verify visual output of Aspose.Cells worksheet rendering in automated tests | detect differences between generated and baseline PNG images in .NET
// Tags: Aspose.Cells worksheet PNG rendering | C# pixel‑by‑pixel image comparison | baseline PNG verification for Excel export | byte array PNG diff in .NET | visual regression testing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook, fills cells, renders the first worksheet to a PNG file using Aspose.Cells, then loads a baseline PNG and performs a byte‑wise pixel comparison, reporting whether the images match and the number of differing pixels.
class Program
{
    static void Main(string[] args)
    {
        try
        {
            // Paths for generated image and baseline reference image
            string generatedImagePath = "generated.png";
            string baselineImagePath = "baseline.png";

            // -------------------------------------------------
            // 1. Create a workbook and populate some data
            // -------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sample";

            // Fill sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);
            sheet.Cells["C2"].PutValue(789);

            // -------------------------------------------------
            // 2. Render the first worksheet to a PNG image
            // -------------------------------------------------
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Default image format is PNG; explicit setting omitted for compatibility
                OnePagePerSheet = true,
                Transparent = false
            };
            SheetRender renderer = new SheetRender(sheet, imgOptions);
            // Render page 0 (the only page) to the specified file
            renderer.ToImage(0, generatedImagePath);

            // -------------------------------------------------
            // 3. Compare the generated image with the baseline
            // -------------------------------------------------
            if (!File.Exists(baselineImagePath))
            {
                Console.WriteLine($"Baseline image not found at path: {baselineImagePath}");
                return;
            }

            int differingPixels;
            bool imagesAreIdentical = CompareImagesPixelByPixel(generatedImagePath, baselineImagePath, out differingPixels);

            if (imagesAreIdentical)
            {
                Console.WriteLine("Images are identical.");
            }
            else
            {
                Console.WriteLine($"Images differ. Number of differing bytes: {differingPixels}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Simple byte‑wise comparison of two image files
    static bool CompareImagesPixelByPixel(string pathA, string pathB, out int diffCount)
    {
        diffCount = 0;

        try
        {
            if (!File.Exists(pathA) || !File.Exists(pathB))
            {
                Console.WriteLine("One of the image files does not exist.");
                return false;
            }

            byte[] bytesA = File.ReadAllBytes(pathA);
            byte[] bytesB = File.ReadAllBytes(pathB);

            if (bytesA.Length != bytesB.Length)
            {
                Console.WriteLine("Image file sizes differ.");
                return false;
            }

            for (int i = 0; i < bytesA.Length; i++)
            {
                if (bytesA[i] != bytesB[i])
                {
                    diffCount++;
                }
            }

            return diffCount == 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during image comparison: {ex.Message}");
            return false;
        }
    }
}
