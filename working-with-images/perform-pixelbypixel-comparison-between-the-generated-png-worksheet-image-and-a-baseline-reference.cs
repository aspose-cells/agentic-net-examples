using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPixelComparison
{
    class Program
    {
        static void Main()
        {
            try
            {
                // ---------- 1. Create a sample workbook ----------
                var workbook = new Workbook();
                var sheet = workbook.Worksheets[0];

                // Populate some sample data
                sheet.Cells["A1"].PutValue("Pixel Comparison Demo");
                sheet.Cells["A2"].PutValue(123);
                sheet.Cells["B2"].PutValue(456);
                sheet.Cells["C3"].PutValue(DateTime.Now);

                // ---------- 2. Render the worksheet to a PNG image ----------
                var renderOptions = new ImageOrPrintOptions(); // defaults to PNG
                var sheetRender = new SheetRender(sheet, renderOptions);

                byte[] generatedImageBytes;
                using (var genStream = new MemoryStream())
                {
                    sheetRender.ToImage(0, genStream);
                    generatedImageBytes = genStream.ToArray();

                    // Optional: save the generated image for visual inspection
                    File.WriteAllBytes("generated.png", generatedImageBytes);
                }

                // ---------- 3. Load the baseline reference image ----------
                const string baselinePath = "baseline.png";
                if (!File.Exists(baselinePath))
                {
                    Console.WriteLine($"Baseline image not found at path: {baselinePath}");
                    return;
                }

                byte[] baselineImageBytes = File.ReadAllBytes(baselinePath);

                // ---------- 4. Perform byte‑by‑byte comparison ----------
                bool imagesAreIdentical = CompareByteArrays(baselineImageBytes, generatedImageBytes);

                // ---------- 5. Output the result ----------
                Console.WriteLine(imagesAreIdentical
                    ? "The generated image matches the baseline reference byte by byte."
                    : "The generated image differs from the baseline reference.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Simple byte array comparison
        private static bool CompareByteArrays(byte[] a, byte[] b)
        {
            if (a == null || b == null) return false;
            if (a.Length != b.Length) return false;

            for (int i = 0; i < a.Length; i++)
            {
                if (a[i] != b[i]) return false;
            }
            return true;
        }
    }
}