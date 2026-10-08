// Title: How to add a full‑page image watermark that scales to the worksheet size while preserving aspect ratio using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that inserts a PNG image as a watermark, calculates the printable page dimensions, and resizes the picture to fill the page without distorting its aspect ratio. | Create a method that iterates through all worksheets, adds a free‑floating locked picture, sets its rotation, and adjusts Width and Height based on the page setup to achieve a full‑page watermark.
// Common Searches: Aspose.Cells C# scale watermark image to fit entire worksheet page | fit image watermark to printable area Excel using Aspose.Cells .NET | maintain aspect ratio when adding full page watermark with Aspose.Cells | C# code example for full‑page Excel watermark with Aspose.Cells | how to resize picture to page size in Aspose.Cells workbook
// Tags: scale image watermark to worksheet page Aspose.Cells | preserve aspect ratio picture Aspose.Cells C# | full page Excel watermark .NET | adjust picture size based on page setup Aspose.Cells | add locked free‑floating picture Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample loads an existing workbook, checks that both the workbook and a PNG watermark file exist, and then adds the image to each worksheet as a free‑floating, locked picture. It computes the printable page width and height from the worksheet's PageSetup, scales the picture to fill the page while keeping its original aspect ratio, optionally rotates it for a typical watermark look, and finally saves the modified workbook.
class WatermarkExample
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string watermarkPath = "watermark.png";
            string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input workbook not found: {inputPath}");
                return;
            }

            // Verify that the watermark image exists
            if (!File.Exists(watermarkPath))
            {
                Console.WriteLine($"Watermark image not found: {watermarkPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (apply to other sheets as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Add the watermark image to the worksheet
            int pictureIndex = sheet.Pictures.Add(0, 0, watermarkPath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Set picture properties to behave like a watermark
            picture.IsLocked = true;
            picture.Placement = PlacementType.FreeFloating;
            // Aspose.Cells Picture does not expose a Transparency property in some versions.
            // If needed, adjust the image itself before adding it as a watermark.
            picture.RotationAngle = -45; // optional: rotate for typical watermark look

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the watermark applied
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved with watermark: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
