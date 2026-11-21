// Title: Compare a rendered worksheet PNG to a baseline image pixel‑by‑pixel using Aspose.Cells in C#
// Common Searches: c# compare generated worksheet png with reference image using Aspose.Cells | how to validate Aspose.Cells SheetRender PNG output against baseline file | byte level comparison of Excel worksheet image in .NET | detect differences between two worksheet PNG files rendered by Aspose.Cells | automated image regression test for Excel sheet rendering in C#
// Tags: Aspose.Cells worksheet PNG rendering validation | pixel‑level image comparison .NET | SheetRender PNG output verification | byte‑array file equality Aspose.Cells | Excel image regression testing C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a workbook, populates sample cells, renders the first worksheet to a PNG using SheetRender, then loads a baseline PNG and performs a byte‑by‑byte comparison to report size mismatches or the first differing byte, indicating whether the images are identical.
class WorksheetImageComparer
{
    static void Main()
    {
        try
        {
            // Paths for generated image and baseline reference image
            string generatedImagePath = "generated.png";
            string baselineImagePath = "baseline.png";

            // -------------------------------------------------
            // 1. Create a workbook and populate some data
            // -------------------------------------------------
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "Sample";

            // Fill some sample data
            ws.Cells["A1"].PutValue("Header");
            ws.Cells["A2"].PutValue(123);
            ws.Cells["B2"].PutValue(456);
            ws.Cells["C2"].PutValue(789);
            ws.Cells["A3"].PutValue(DateTime.Now);
            ws.Cells["B3"].PutValue(true);

            // -------------------------------------------------
            // 2. Render the worksheet to a PNG image
            // -------------------------------------------------
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // ImageFormat defaults to PNG, so no need to set it explicitly
                OnePagePerSheet = true,
                Transparent = false,
                HorizontalResolution = 96,
                VerticalResolution = 96
            };

            // Render the first worksheet (index 0)
            SheetRender sr = new SheetRender(ws, imgOptions);
            sr.ToImage(0, generatedImagePath); // Save as PNG

            // -------------------------------------------------
            // 3. Verify baseline image exists before comparison
            // -------------------------------------------------
            if (!File.Exists(baselineImagePath))
            {
                Console.WriteLine($"Baseline image not found at '{baselineImagePath}'. Comparison skipped.");
                return;
            }

            // -------------------------------------------------
            // 4. Load both images for byte‑by‑byte comparison
            // -------------------------------------------------
            bool imagesAreEqual = CompareFilesByteByByte(generatedImagePath, baselineImagePath, out string diffInfo);
            Console.WriteLine(imagesAreEqual
                ? "Images are identical."
                : $"Images differ: {diffInfo}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    private static bool CompareFilesByteByByte(string filePath1, string filePath2, out string diffInfo)
    {
        diffInfo = string.Empty;

        try
        {
            byte[] bytes1 = File.ReadAllBytes(filePath1);
            byte[] bytes2 = File.ReadAllBytes(filePath2);

            if (bytes1.Length != bytes2.Length)
            {
                diffInfo = $"File size mismatch. Generated: {bytes1.Length} bytes, Baseline: {bytes2.Length} bytes.";
                return false;
            }

            for (int i = 0; i < bytes1.Length; i++)
            {
                if (bytes1[i] != bytes2[i])
                {
                    diffInfo = $"Byte difference at position {i}. Generated=0x{bytes1[i]:X2}, Baseline=0x{bytes2[i]:X2}.";
                    return false;
                }
            }

            return true;
        }
        catch (Exception ex)
        {
            diffInfo = $"Comparison failed: {ex.Message}";
            return false;
        }
    }
}
