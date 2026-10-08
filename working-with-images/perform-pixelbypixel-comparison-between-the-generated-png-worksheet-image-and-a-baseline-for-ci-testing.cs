// Title: Render an Excel worksheet to PNG with Aspose.Cells and verify pixel‑level consistency against a baseline image in CI
// AI Prompts: Use Aspose.Cells to render the first worksheet of a workbook to a PNG stream, load a stored baseline PNG file, compare the two byte arrays element‑by‑element, and output the count of mismatched bytes. | Build a .NET console utility for continuous integration that generates a worksheet image, creates a baseline PNG when missing, performs a byte‑wise PNG comparison, writes the result to the console, and exits with a non‑zero code if differences are detected.
// Common Searches: how to compare generated worksheet png with a reference image using Aspose.Cells in C# | ci pipeline test for Excel sheet rendering to PNG .NET | byte array diff count for two png files in a C# console application | automated regression testing of Excel worksheet images with Aspose.Cells | generate baseline image for Excel rendering and compare in continuous integration
// Tags: Aspose.Cells worksheet to PNG rendering | pixel‑level PNG comparison in C# | CI regression testing of Excel image output | byte‑wise image diff for Aspose.Cells | baseline image generation for Excel rendering

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example renders the first worksheet of a workbook to a PNG using Aspose.Cells, loads (or creates) a baseline PNG, performs a byte‑by‑byte comparison to detect any differences, reports the mismatch count, and exits with an appropriate status code for CI pipelines.
public class WorksheetImageComparer
{
    // Generates a PNG image of the first worksheet and returns it as a byte array
    private static byte[] GenerateWorksheetImage()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data (replace with actual test data as needed)
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);
            sheet.Cells["C2"].PutValue(789);

            // Define image options for rendering the worksheet
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true,
                Transparent = false,
                HorizontalResolution = 96,
                VerticalResolution = 96
            };

            // Render the worksheet to an image stored in a memory stream
            using (MemoryStream ms = new MemoryStream())
            {
                SheetRender sheetRender = new SheetRender(sheet, imgOptions);
                sheetRender.ToImage(0, ms);
                return ms.ToArray(); // Return PNG bytes
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("Error generating worksheet image: " + ex.Message);
            throw;
        }
    }

    // Loads a baseline PNG image from disk and returns its bytes
    private static byte[] LoadBaselineImage(string baselinePath)
    {
        if (!File.Exists(baselinePath))
            throw new FileNotFoundException("Baseline image not found.", baselinePath);

        return File.ReadAllBytes(baselinePath);
    }

    // Performs a byte‑by‑byte comparison between two PNG byte arrays
    // Returns true if images are identical; diffCount receives the number of differing bytes
    public static bool CompareByteArrays(byte[] data1, byte[] data2, out int diffCount)
    {
        diffCount = 0;

        // Quick size check
        if (data1.Length != data2.Length)
            return false;

        for (int i = 0; i < data1.Length; i++)
        {
            if (data1[i] != data2[i])
                diffCount++;
        }

        return diffCount == 0;
    }

    // Entry point for CI test
    public static void Main(string[] args)
    {
        try
        {
            string baselinePath = @"baseline.png";

            // Generate the worksheet image
            byte[] generatedImage = GenerateWorksheetImage();

            // Ensure baseline exists; if not, create it from the generated image
            if (!File.Exists(baselinePath))
            {
                File.WriteAllBytes(baselinePath, generatedImage);
                Console.WriteLine("Baseline image not found. Created new baseline at: " + baselinePath);
            }

            // Load baseline image
            byte[] baselineImage = LoadBaselineImage(baselinePath);

            // Compare images
            int differingBytes;
            bool areIdentical = CompareByteArrays(generatedImage, baselineImage, out differingBytes);

            // Output result (CI system can capture console output)
            Console.WriteLine("Images identical: {0}", areIdentical);
            Console.WriteLine("Differing bytes: {0}", differingBytes);

            // Optionally, save the generated image for debugging
            string debugPath = @"generated_debug.png";
            File.WriteAllBytes(debugPath, generatedImage);
            Console.WriteLine("Generated image saved to: " + debugPath);

            // Exit code can be used by CI to indicate success/failure
            Environment.Exit(areIdentical ? 0 : 1);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("Error: " + ex.Message);
            Environment.Exit(1);
        }
    }
}
