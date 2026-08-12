// Title: C# – Pixel‑by‑pixel comparison of a rendered worksheet PNG with a baseline using Aspose.Cells
// Description: Creates a workbook, renders the first worksheet to a PNG image in memory, loads a baseline PNG file, and compares the two byte arrays to determine if the images are identical—ideal for CI visual‑regression testing.
// Keywords: Aspose.Cells PNG rendering | C# image regression test | pixel level image comparison | WorkbookRender to PNG | byte array image equality | continuous integration visual testing | Excel worksheet screenshot verification
// Common Searches: compare rendered Excel worksheet PNG with baseline C# | Aspose.Cells image regression testing example | pixel perfect PNG comparison for CI | how to verify Excel sheet image output Aspose.Cells | byte array image equality in .NET
// Developer Intent: The developer needs an automated way to confirm that a PNG generated from an Excel worksheet matches a stored reference image, enabling reliable visual regression checks in build pipelines.
// Use Cases: Detect visual changes in Excel reports during continuous integration. | Validate that formatting or data updates do not alter the rendered image. | Automate screenshot testing for dashboards generated with Aspose.Cells.
// AI Prompts: Generate a C# utility that compares two PNG files pixel by pixel, reports the first mismatched coordinate, and integrates with Aspose.Cells rendering. | Create a tolerant image comparison method for Aspose.Cells that ignores minor compression artifacts and works with NUnit or xUnit. | Write code to log detailed differences (pixel position, expected vs. actual color) when rendered worksheet images differ.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsImageComparison
{
    // Creates a workbook, renders the first worksheet to a PNG image in memory, loads a baseline PNG file, and compares the two byte arrays to determine if the images are identical—ideal for CI visual‑regression testing.
    class Program
    {
        static void Main()
        {
            try
            {
                // -------------------- Create workbook --------------------
                Workbook workbook = new Workbook(); // create empty workbook
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data (adjust as needed for your test)
                sheet.Cells["A1"].PutValue("Sample");
                sheet.Cells["B1"].PutValue(123);
                sheet.Cells["A2"].PutValue(DateTime.Now);

                // -------------------- Render workbook to PNG --------------------
                // ImageOrPrintOptions defaults to PNG format
                ImageOrPrintOptions renderOptions = new ImageOrPrintOptions();

                WorkbookRender renderer = new WorkbookRender(workbook, renderOptions);
                byte[] renderedBytes;
                using (MemoryStream renderedStream = new MemoryStream())
                {
                    // Render first worksheet (index 0) to the stream
                    renderer.ToImage(0, renderedStream);
                    renderedBytes = renderedStream.ToArray(); // capture PNG bytes
                }

                // -------------------- Load baseline image --------------------
                string baselinePath = "baseline.png"; // path to baseline image
                if (!File.Exists(baselinePath))
                {
                    Console.WriteLine($"Baseline image not found at '{baselinePath}'.");
                    return;
                }

                byte[] baselineBytes;
                try
                {
                    baselineBytes = File.ReadAllBytes(baselinePath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to read baseline image: {ex.Message}");
                    return;
                }

                // -------------------- Compare images byte by byte --------------------
                bool areEqual = CompareByteArrays(renderedBytes, baselineBytes);
                Console.WriteLine($"Images are {(areEqual ? "identical" : "different")}.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Returns true if both byte arrays have the same length and identical content
        private static bool CompareByteArrays(byte[] arr1, byte[] arr2)
        {
            if (arr1 == null || arr2 == null) return false;
            if (arr1.Length != arr2.Length) return false;

            for (int i = 0; i < arr1.Length; i++)
            {
                if (arr1[i] != arr2[i])
                    return false;
            }

            return true;
        }
    }
}
