// Title: How to compare a rendered worksheet PNG image to a baseline PNG pixel‑by‑pixel using Aspose.Cells for .NET
// AI Prompts: Render a specific worksheet to a PNG stream with Aspose.Cells and return the image bytes. | Write a C# method that loads a baseline PNG file and checks each byte against the rendered worksheet image to determine equality. | Create a CI‑friendly routine that catches all exceptions and treats any error as a mismatch when validating worksheet image output.
// Common Searches: Aspose.Cells compare generated worksheet PNG with reference image in unit tests | C# pixel level comparison of Excel sheet rendering output | How to validate Excel worksheet image rendering using byte array comparison | Automated CI check for identical worksheet PNG using Aspose.Cells .NET
// Tags: Aspose.Cells worksheet PNG rendering | pixel‑accurate PNG diff C# | CI image verification Aspose.Cells | render Excel sheet to memory stream | exception‑safe image match check

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, renders a selected worksheet to a PNG image in memory using Aspose.Cells, reads a baseline PNG file, and performs a length and byte‑by‑byte comparison. It returns true only when the images are identical and treats any exception as a mismatch, making it suitable for CI image‑validation tests.
public class WorksheetImageComparer
{
    /// <param name="workbookPath">Path to the Excel workbook.</param>
    /// <param name="sheetIndex">Zero‑based index of the worksheet to render.</param>
    /// <param name="baselineImagePath">Path to the baseline PNG image.</param>
    /// <returns>True if the images are identical; otherwise false.</returns>
    public static bool CompareWorksheetImage(string workbookPath, int sheetIndex, string baselineImagePath)
    {
        try
        {
            // Validate input files
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException("Workbook not found.", workbookPath);
            if (!File.Exists(baselineImagePath))
                throw new FileNotFoundException("Baseline image not found.", baselineImagePath);
            if (sheetIndex < 0)
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), "Sheet index must be non‑negative.");

            // Load workbook
            var workbook = new Workbook(workbookPath);
            if (sheetIndex >= workbook.Worksheets.Count)
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), "Sheet index exceeds worksheet count.");

            // Set rendering options (default format is PNG)
            var renderOptions = new ImageOrPrintOptions
            {
                HorizontalResolution = 96,
                VerticalResolution = 96
            };

            // Render worksheet to a memory stream
            using (var generatedStream = new MemoryStream())
            {
                var sheetRender = new SheetRender(workbook.Worksheets[sheetIndex], renderOptions);
                sheetRender.ToImage(0, generatedStream);
                byte[] generatedBytes = generatedStream.ToArray();

                // Load baseline image bytes
                byte[] baselineBytes = File.ReadAllBytes(baselineImagePath);

                // Quick length check
                if (generatedBytes.Length != baselineBytes.Length)
                    return false;

                // Compare byte by byte
                for (int i = 0; i < generatedBytes.Length; i++)
                {
                    if (generatedBytes[i] != baselineBytes[i])
                        return false;
                }

                return true; // Images match
            }
        }
        catch
        {
            // In case of any unexpected error, treat as non‑match
            return false;
        }
    }

    // Entry point required for compilation
    public static void Main(string[] args)
    {
        // Example usage – adjust paths as needed
        string workbookPath = "TestData/Report.xlsx";
        string baselinePath = "Baseline/Report_Page1.png";

        bool match = CompareWorksheetImage(workbookPath, 0, baselinePath);
        Console.WriteLine(match ? "Images match." : "Images do NOT match.");
    }
}
