// Title: Add a reusable WordArt watermark with custom text, style, and lock option using Aspose.Cells for .NET
// Description: C# example that defines a static helper method to insert a WordArt shape into a chosen worksheet, apply a preset WordArt style, set font attributes, and optionally lock the text so it stays unchanged when the sheet is protected, enabling consistent watermarks across workbooks.
// Keywords: Aspose.Cells | WordArt watermark | C# | .NET | preset WordArt style | lock text | worksheet protection | reusable method | Excel watermark | AddWordArtWatermark
// Common Searches: Aspose.Cells add WordArt watermark C# | lock WordArt text in protected sheet Aspose.Cells | reusable watermark method Aspose.Cells .NET | preset WordArt styles Excel watermark | how to insert WordArt shape with Aspose.Cells
// Developer Intent: Insert a WordArt watermark with specified text, chosen style, and optional lock into a worksheet using Aspose.Cells.
// Use Cases: Generate confidential reports by adding a locked "CONFIDENTIAL" WordArt watermark to the first sheet. | Apply brand‑specific WordArt watermarks across multiple worksheets with a single helper method. | Automate watermark insertion in batch processing while respecting sheet protection settings. | Create dynamic watermarks (e.g., date, user name) in generated Excel files.
// AI Prompts: Write a C# function using Aspose.Cells that adds a WordArt watermark with custom text, a selected PresetWordArtStyle, and a boolean to lock the text on protected sheets. | Extend AddWordArtWatermark to accept font size, bold, and italic parameters and return the created Shape object. | Provide sample code that iterates through all worksheets in a workbook and calls AddWordArtWatermark with different texts and styles. | Show how to combine AddWordArtWatermark with workbook.Save in a web API that returns watermarked Excel files.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// C# example that defines a static helper method to insert a WordArt shape into a chosen worksheet, apply a preset WordArt style, set font attributes, and optionally lock the text so it stays unchanged when the sheet is protected, enabling consistent watermarks across workbooks.
public static class WatermarkHelper
{
    // Adds a WordArt watermark to the specified worksheet.
    // Parameters:
    //   workbook   - target workbook
    //   sheetIndex - index of the worksheet to receive the watermark
    //   text       - watermark text
    //   style      - preset WordArt style
    //   lockText   - whether the text should be locked when the sheet is protected
    public static void AddWordArtWatermark(Workbook workbook, int sheetIndex, string text, PresetWordArtStyle style, bool lockText)
    {
        Worksheet ws = workbook.Worksheets[sheetIndex];
        ShapeCollection shapes = ws.Shapes;

        // Position and size of the WordArt (adjust as needed)
        int topRow = 0;
        int top = 0;
        int leftColumn = 0;
        int left = 0;
        int height = 200; // pixels
        int width = 600;  // pixels

        // Insert the WordArt shape
        Shape wordArt = shapes.AddWordArt(style, text, topRow, top, leftColumn, left, height, width);

        // Customize appearance via TextEffect
        TextEffectFormat txtEffect = wordArt.TextEffect;
        txtEffect.FontBold = true;
        txtEffect.FontItalic = false;
        txtEffect.FontSize = 48;

        // Apply lock setting to the text
        var textAlignment = wordArt.TextBody.TextAlignment;
        textAlignment.IsLockedText = lockText;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // Add a WordArt watermark to the first worksheet
        WatermarkHelper.AddWordArtWatermark(
            wb,
            0,
            "CONFIDENTIAL",
            PresetWordArtStyle.WordArtStyle7,
            true);

        // Save the workbook
        wb.Save("WatermarkDemo.xlsx");
    }
}
