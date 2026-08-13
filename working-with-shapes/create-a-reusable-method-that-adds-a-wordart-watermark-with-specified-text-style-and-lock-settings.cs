// Title: C# helper to insert a WordArt watermark with custom text, style, transparency, rotation, and lock option using Aspose.Cells
// Description: Provides a static AddWordArtWatermark method that creates a WordArt shape on a given worksheet, applies a preset style, sets 70 % transparency, rotates it 45°, and can lock the text for protected sheets. The sample program builds a workbook, adds the watermark, and writes the file.
// Keywords: Aspose.Cells | C# WordArt watermark | Excel watermark Aspose | preset WordArt style | watermark transparency | rotate WordArt | lock watermark text | worksheet protection | AddWordArtWatermark | shape collection | .NET Excel shaping
// Common Searches: Add WordArt watermark with Aspose.Cells C# | Set transparency for WordArt shape in Excel using Aspose | Rotate WordArt watermark programmatically | Lock WordArt text when protecting worksheet Aspose.Cells | Reusable method for Excel watermarks .NET
// Developer Intent: Add a semi‑transparent, rotated WordArt shape as a protected watermark to a chosen worksheet.
// Use Cases: Mark confidential reports with a diagonal “CONFIDENTIAL” watermark before distribution. | Apply corporate branding WordArt to every sheet in a generated workbook. | Prevent users from editing the watermark by locking the text after sheet protection. | Create a template that automatically inserts a styled watermark on new worksheets.
// AI Prompts: Write C# code that uses Aspose.Cells to add a WordArt watermark with custom text, preset style, 70 % transparency, 45° rotation, and an optional lock flag. | Explain how to modify AddWordArtWatermark to accept custom top‑row, left‑column, height, and width parameters. | Show how to protect a worksheet after inserting a locked WordArt watermark with Aspose.Cells. | Provide a unit‑test example for the AddWordArtWatermark method. | Generate XML documentation comments for the AddWordArtWatermark helper.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// Provides a static AddWordArtWatermark method that creates a WordArt shape on a given worksheet, applies a preset style, sets 70 % transparency, rotates it 45°, and can lock the text for protected sheets. The sample program builds a workbook, adds the watermark, and writes the file.
public static class WatermarkHelper
{
    /// <param name="workbook">The workbook to modify.</param>
    /// <param name="sheetName">Name of the worksheet where the watermark will be placed.</param>
    /// <param name="text">Watermark text.</param>
    /// <param name="style">Preset WordArt style.</param>
    /// <param name="lockText">If true, the watermark text is locked when the sheet is protected.</param>
    public static void AddWordArtWatermark(Workbook workbook, string sheetName, string text, PresetWordArtStyle style, bool lockText)
    {
        // Get the target worksheet
        Worksheet sheet = workbook.Worksheets[sheetName];

        // Access the shape collection of the worksheet
        ShapeCollection shapes = sheet.Shapes;

        // Add a WordArt shape.
        // Parameters: style, text, topRow, top, leftColumn, left, height, width
        // Position it at the top‑left corner and give it a reasonable size.
        Shape wordArt = shapes.AddWordArt(
            style,
            text,
            topRow: 0,
            top: 0,
            leftColumn: 0,
            left: 0,
            height: 100,
            width: 400);

        // Make the WordArt semi‑transparent (use the current Fill property)
        if (wordArt.Fill != null)
        {
            wordArt.Fill.Transparency = 0.7; // 70% transparent
        }

        // Rotate the shape to achieve a typical watermark appearance
        wordArt.RotationAngle = 45; // degrees

        // Lock the text if required (effective when the worksheet is protected)
        if (wordArt.TextBody != null && wordArt.TextBody.TextAlignment != null)
        {
            wordArt.TextBody.TextAlignment.IsLockedText = lockText;
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
            // Create a new workbook and add some sample data
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Sample data for the worksheet.");

            // Add a WordArt watermark
            WatermarkHelper.AddWordArtWatermark(
                workbook: wb,
                sheetName: ws.Name,
                text: "CONFIDENTIAL",
                style: PresetWordArtStyle.WordArtStyle7,
                lockText: true);

            // Save the workbook
            string outputPath = "WordArtWatermark.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
