using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

public static class WatermarkHelper
{
    /// <summary>
    /// Adds a WordArt watermark to the specified worksheet.
    /// </summary>
    /// <param name="workbook">The workbook to modify.</param>
    /// <param name="sheetName">Name of the worksheet where the watermark will be placed.</param>
    /// <param name="text">Watermark text.</param>
    /// <param name="style">Preset WordArt style.</param>
    /// <param name="lockText">If true, the watermark text is locked when the sheet is protected.</param>
    public static void AddWordArtWatermark(Workbook workbook, string sheetName, string text, PresetWordArtStyle style, bool lockText)
    {
        try
        {
            // Ensure the target worksheet exists; create it if missing
            Worksheet sheet = workbook.Worksheets[sheetName] ?? workbook.Worksheets.Add(sheetName);

            // Add a WordArt shape with default position and size
            Shape wordArt = sheet.Shapes.AddWordArt(
                style,
                text,
                0,          // topRow
                0,          // top (pixel offset)
                0,          // leftColumn
                0,          // left (pixel offset)
                100,        // height (pixels)
                400);       // width (pixels)

            // Lock the text of the shape if required
            wordArt.TextBody.TextAlignment.IsLockedText = lockText;

            // Optional: send the shape to the background.
            // Aspose.Cells does not expose a ToBack method; setting ZOrderPosition to a low value achieves a similar effect.
            wordArt.ZOrderPosition = -1;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"AddWordArtWatermark error: {ex.Message}");
            throw;
        }
    }
}

// Example usage
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook wb = new Workbook();

            // Define the worksheet name
            string sheetName = "Sheet1";

            // Add a WordArt watermark with style 2 and lock the text
            WatermarkHelper.AddWordArtWatermark(
                wb,
                sheetName,
                "CONFIDENTIAL",
                PresetWordArtStyle.WordArtStyle2,
                lockText: true);

            // Save the workbook to verify the watermark
            string outputPath = "WordArtWatermark.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}