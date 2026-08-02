// Title: C# – Pixel‑by‑Pixel PNG Comparison of a Rendered Worksheet for CI with Aspose.Cells
// Description: Demonstrates how to render the first worksheet of a workbook to PNG using Aspose.Cells, load a baseline image, compare the two files byte‑by‑byte, and output a PASS/FAIL result suitable for continuous‑integration pipelines.
// Keywords: Aspose.Cells | C# | PNG image comparison | CI testing | continuous integration | byte array diff | WorkbookRender | automated image validation | pixel level diff | GitHub example
// Common Searches: compare rendered worksheet png to baseline c# | aspacells image diff for ci pipelines | bytewise png comparison .net | automated excel sheet image validation | continuous integration image testing aspocells
// Developer Intent: Verify that a PNG generated from an Excel worksheet matches an expected baseline image in an automated CI build.
// Use Cases: Render a worksheet to PNG, read both generated and baseline files as byte arrays, and detect any mismatches. | Integrate the comparison logic into a build script so the pipeline fails when visual output changes. | Run the check after library upgrades to ensure rendering consistency.
// AI Prompts: Create a C# function that compares two PNG files pixel by pixel using System.Drawing.Bitmap for visual diff. | Show how to add Aspose.Cells image rendering and byte‑wise comparison to an Azure DevOps pipeline task that aborts on failures. | Generate an NUnit test that asserts the rendered worksheet image equals a stored baseline and reports detailed differences.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsCiTesting
{
    // Demonstrates how to render the first worksheet of a workbook to PNG using Aspose.Cells, load a baseline image, compare the two files byte‑by‑byte, and output a PASS/FAIL result suitable for continuous‑integration pipelines.
    public class WorksheetImageComparison
    {
        // Entry point for the console application
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Core demo logic
        public static void Run()
        {
            // -----------------------------------------------------------------
            // 1. Create a sample workbook and populate it with some data
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells CI Image Comparison");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B1"].PutValue(123);
            sheet.Cells["B2"].PutValue(456.78);

            // -----------------------------------------------------------------
            // 2. Render the first worksheet page to a PNG image file
            //    Using the WorkbookRender.ToImage(string) overload.
            // -----------------------------------------------------------------
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                // PNG is inferred from the file extension; explicit setting omitted for compatibility
                OnePagePerSheet = true
            };
            WorkbookRender renderer = new WorkbookRender(workbook, renderOptions);
            string generatedImagePath = Path.Combine("output", "generated.png");
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(generatedImagePath) ?? string.Empty);
                renderer.ToImage(generatedImagePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error rendering image: {ex.Message}");
                return;
            }

            // -----------------------------------------------------------------
            // 3. Load the baseline image that represents the expected output.
            // -----------------------------------------------------------------
            string baselineImagePath = Path.Combine("baseline", "expected.png");
            if (!File.Exists(baselineImagePath))
            {
                Console.WriteLine($"Baseline image not found at: {baselineImagePath}");
                return;
            }

            // -----------------------------------------------------------------
            // 4. Perform byte‑by‑byte comparison of the two PNG files.
            // -----------------------------------------------------------------
            if (!File.Exists(generatedImagePath))
            {
                Console.WriteLine($"Generated image not found at: {generatedImagePath}");
                return;
            }

            byte[] generatedBytes;
            byte[] baselineBytes;
            try
            {
                generatedBytes = File.ReadAllBytes(generatedImagePath);
                baselineBytes = File.ReadAllBytes(baselineImagePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading image files: {ex.Message}");
                return;
            }

            bool imagesAreIdentical = generatedBytes.Length == baselineBytes.Length;
            int diffCount = 0;

            if (imagesAreIdentical)
            {
                for (int i = 0; i < generatedBytes.Length; i++)
                {
                    if (generatedBytes[i] != baselineBytes[i])
                    {
                        imagesAreIdentical = false;
                        diffCount++;
                    }
                }
            }
            else
            {
                diffCount = Math.Abs(generatedBytes.Length - baselineBytes.Length);
                Console.WriteLine("Image file sizes differ.");
            }

            // -----------------------------------------------------------------
            // 5. Report the result – suitable for CI pipelines.
            // -----------------------------------------------------------------
            if (imagesAreIdentical)
            {
                Console.WriteLine("PASS: Generated image matches the baseline.");
            }
            else
            {
                Console.WriteLine($"FAIL: Images differ. Byte mismatches: {diffCount}");
            }
        }
    }
}
