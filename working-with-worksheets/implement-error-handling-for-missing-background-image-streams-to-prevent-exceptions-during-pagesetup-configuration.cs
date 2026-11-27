// Title: Add a background image to an Aspose.Cells worksheet in C# with safe handling for missing files
// Common Searches: C# Aspose.Cells add worksheet background image only if file exists | How to prevent exception when background picture file is missing in Aspose.Cells | Aspose.Cells picture.Add error handling for missing PNG in .NET | Safely set Excel worksheet background using Aspose.Cells without crashing | Check image file before adding to Aspose.Cells worksheet in C#
// Tags: Aspose.Cells worksheet background image addition | C# file existence check for Aspose.Cells picture | exception handling for Aspose.Cells picture.Add | prevent missing image stream error Aspose.Cells | safe Excel background insertion using Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a Workbook, accesses the first Worksheet, verifies that a PNG file exists, and then adds it as a picture anchored at cell A1 inside a try‑catch block. If the file is missing or the addition fails, a friendly message is logged and execution continues. The output directory is ensured before saving the workbook to 'output.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the path to the background image
            string backgroundImagePath = "background.png";

            // Verify that the background image file exists
            if (File.Exists(backgroundImagePath))
            {
                try
                {
                    // Add the image as a picture anchored at cell A1 (row 0, column 0)
                    sheet.Pictures.Add(0, 0, backgroundImagePath);
                }
                catch (Exception imgEx)
                {
                    Console.WriteLine($"Failed to add background image: {imgEx.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Background image file '{backgroundImagePath}' not found. Skipping background setup.");
            }

            // Define output path and ensure the directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
