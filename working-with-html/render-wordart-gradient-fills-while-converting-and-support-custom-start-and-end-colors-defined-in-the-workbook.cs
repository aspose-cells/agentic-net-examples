// Title: Create WordArt with a custom two‑color gradient from worksheet RGB values and save as PDF using Aspose.Cells for .NET
// Common Searches: how to set a two‑color gradient fill on WordArt shape using Aspose.Cells C# | read RGB color strings from Excel cells and apply to WordArt gradient in Aspose.Cells | export WordArt with vertical gradient to PDF with Aspose.Cells for .NET | parse R,G,B color values from worksheet and use in Aspose.Cells shape fill | apply custom start and end colors to WordArt gradient in Aspose.Cells
// Tags: WordArt two-color gradient Aspose.Cells | parse RGB string to System.Drawing.Color in C# | apply gradient fill to WordArt shape | save workbook as PDF with gradient WordArt | PresetWordArtStyle configuration Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example stores start and end RGB values as strings in cells A1 and A2, parses them into System.Drawing.Color objects, adds a WordArt shape, applies a vertical two‑color gradient using the parsed colors, sets font styling for visibility, and saves the workbook as a PDF, demonstrating custom gradient fills for WordArt with Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // ------------------------------------------------------------
        // Define custom start and end colors for the gradient in cells.
        // The colors are stored as "R,G,B" strings (e.g., "255,0,0").
        // ------------------------------------------------------------
        worksheet.Cells["A1"].PutValue("255,0,0"); // Start color – Red
        worksheet.Cells["A2"].PutValue("0,0,255"); // End color – Blue

        // Parse the color strings from the cells
        Color startColor = ParseColor(worksheet.Cells["A1"].StringValue);
        Color endColor = ParseColor(worksheet.Cells["A2"].StringValue);

        // ------------------------------------------------------------
        // Add a WordArt shape to the worksheet.
        // ------------------------------------------------------------
        // Parameters: style, text, topRow, top (pixel offset), leftColumn, left (pixel offset), height, width
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1,
            "Gradient WordArt",
            2,   // topRow
            10,  // top offset (pixels)
            2,   // leftColumn
            10,  // left offset (pixels)
            100, // height (pixels)
            300  // width (pixels)
        );

        // ------------------------------------------------------------
        // Apply a two‑color gradient fill to the WordArt.
        // ------------------------------------------------------------
        wordArt.Fill.FillType = FillType.Gradient;
        // GradientStyleType.Vertical creates a top‑to‑bottom gradient.
        // Variant = 1 corresponds to the first variant in the Fill Effects dialog.
        wordArt.Fill.SetTwoColorGradient(startColor, endColor, GradientStyleType.Vertical, 1);

        // Optional: adjust font properties so the text is visible over the gradient.
        wordArt.Font.Size = 24;
        wordArt.Font.Color = Color.White;
        wordArt.Font.IsBold = true;

        // ------------------------------------------------------------
        // Save the workbook to PDF (or any other format) to render the gradient.
        // ------------------------------------------------------------
        string outputPath = "WordArtGradient.pdf";
        workbook.Save(outputPath, SaveFormat.Pdf);
        Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
    }

    static Color ParseColor(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return Color.Empty;

        value = value.Trim();

        // HTML hex notation
        if (value.StartsWith("#"))
        {
            try
            {
                return ColorTranslator.FromHtml(value);
            }
            catch
            {
                return Color.Empty;
            }
        }

        // "R,G,B" notation
        string[] parts = value.Split(',');
        if (parts.Length == 3 &&
            int.TryParse(parts[0], out int r) &&
            int.TryParse(parts[1], out int g) &&
            int.TryParse(parts[2], out int b))
        {
            // Clamp values to 0‑255 just in case
            r = Math.Clamp(r, 0, 255);
            g = Math.Clamp(g, 0, 255);
            b = Math.Clamp(b, 0, 255);
            return Color.FromArgb(r, g, b);
        }

        return Color.Empty;
    }
}
