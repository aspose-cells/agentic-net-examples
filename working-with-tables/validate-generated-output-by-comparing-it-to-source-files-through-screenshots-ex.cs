using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsValidation
{
    class Program
    {
        // Compare two byte arrays.
        static bool CompareBytes(byte[] a, byte[] b)
        {
            if (a == null || b == null) return false;
            if (a.Length != b.Length) return false;
            for (int i = 0; i < a.Length; i++)
                if (a[i] != b[i]) return false;
            return true;
        }

        static void Main()
        {
            // Paths – adjust as needed.
            string sourceWorkbookPath = "source.xlsx";                 // Original workbook to validate.
            string referenceScreenshotsFolder = "ReferenceScreenshots"; // Folder containing expected page images.
            string renderedScreenshotsFolder = "RenderedScreenshots";   // Folder to store generated images.

            // Ensure output folder exists.
            Directory.CreateDirectory(renderedScreenshotsFolder);

            // Load the source workbook.
            Workbook workbook = new Workbook(sourceWorkbookPath);

            // Configure image rendering options (PNG format).
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                OnePagePerSheet = true
            };

            // Create the renderer.
            WorkbookRender renderer = new WorkbookRender(workbook, renderOptions);

            // Iterate through each rendered page.
            for (int pageIndex = 0; pageIndex < renderer.PageCount; pageIndex++)
            {
                // Render page to a memory stream.
                using (MemoryStream renderedStream = new MemoryStream())
                {
                    renderer.ToImage(pageIndex, renderedStream);
                    byte[] renderedBytes = renderedStream.ToArray();

                    // Save rendered image for visual inspection (optional).
                    string renderedPath = Path.Combine(renderedScreenshotsFolder, $"page_{pageIndex}.png");
                    File.WriteAllBytes(renderedPath, renderedBytes);

                    // Load reference screenshot.
                    string referencePath = Path.Combine(referenceScreenshotsFolder, $"page_{pageIndex}.png");
                    if (!File.Exists(referencePath))
                    {
                        Console.WriteLine($"Reference image not found for page {pageIndex}: {referencePath}");
                        continue;
                    }

                    byte[] referenceBytes = File.ReadAllBytes(referencePath);

                    // Compare the two images.
                    bool areEqual = CompareBytes(renderedBytes, referenceBytes);
                    Console.WriteLine($"Page {pageIndex}: {(areEqual ? "MATCH" : "DIFFERENT")}");
                }
            }

            // Clean up.
            renderer.Dispose();
            workbook.Dispose();
        }
    }
}