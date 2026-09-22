// Title: Add a centered Arial 30‑point text watermark with 50% opacity to an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# using Aspose.Cells to place a WordArt shape with Arial 30 pt text as a semi‑transparent watermark on the first worksheet. | Show how to size the watermark shape to the worksheet page dimensions and rotate it -45° for a diagonal effect. | Demonstrate modifying the watermark’s opacity, font, or rotation without affecting other worksheet content.
// Common Searches: Aspose.Cells C# add semi transparent text watermark to Excel worksheet | fit watermark shape to entire page size in Aspose.Cells workbook | rotate Aspose.Cells text effect -45 degrees for diagonal watermark | change watermark opacity to 50% using Aspose.Cells shape transparency
// Tags: add semi transparent text watermark Aspose.Cells | fit watermark shape to worksheet page Aspose.Cells | apply diagonal rotation to text effect Aspose.Cells | set shape fill transparency Aspose.Cells | use Arial 30pt font for Excel watermark Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// // Loads an existing workbook, creates a WordArt shape with Arial 30 pt text, sizes it to the page, sets 50 % opacity, rotates -45°, sends it to the back, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a text shape (WordArt) to act as a watermark
            // Parameters: preset effect, text, font name, font size, bold, italic,
            // left, top, width, height, anchor row, anchor column
            Shape watermarkShape = worksheet.Shapes.AddTextEffect(
                MsoPresetTextEffect.TextEffect1,
                "Sample Watermark",
                "Arial",
                30,
                false,
                false,
                0,
                0,
                500,
                100,
                0,
                0);

            // Position the shape to cover the whole page
            watermarkShape.Placement = PlacementType.FreeFloating;
            watermarkShape.Left = 0;
            watermarkShape.Top = 0;

            // Convert paper size from inches to points (1 inch = 72 points) and cast to int
            watermarkShape.Width = (int)(worksheet.PageSetup.PaperWidth * 72);
            watermarkShape.Height = (int)(worksheet.PageSetup.PaperHeight * 72);

            // Set transparency (0 = opaque, 1 = fully transparent)
            watermarkShape.Fill.Transparency = 0.5; // 50% opacity

            // Rotate the watermark for diagonal appearance
            watermarkShape.RotationAngle = -45;

            // Send the shape to the back by setting the lowest Z-order position
            watermarkShape.ZOrderPosition = 0;

            // Save the workbook with the watermark applied
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
