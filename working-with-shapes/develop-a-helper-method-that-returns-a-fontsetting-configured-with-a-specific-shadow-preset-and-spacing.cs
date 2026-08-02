// Title: Aspose.Cells .NET: Apply a Preset Shadow and Custom Spacing to Shape Text
// Description: A C# helper that retrieves a FontSetting for a specified character range in a Shape, assigns a PresetShadowType, sets character spacing, and returns the configured object for further font styling before saving the workbook.
// Keywords: Aspose.Cells FontSetting shadow | preset shadow type C# | character spacing shape text | Aspose.Cells shape text formatting | C# Aspose.Cells FontSetting example | apply text shadow Aspose.Cells | shape text character range | Aspose.Cells .NET tutorial
// Common Searches: how to set a preset shadow on shape text using Aspose.Cells | change character spacing for shape text in C# Aspose.Cells | Aspose.Cells FontSetting shadow and spacing example | apply text effects to a shape in Aspose.Cells .NET | C# code to format shape text with shadow and spacing
// Developer Intent: Create a FontSetting that applies a chosen preset shadow and a specific character spacing to a defined range of characters inside a Shape.
// Use Cases: Add an offset‑bottom shadow and increase spacing for the entire text of a rectangle shape. | Apply a tight spacing and a perspective shadow only to the first five characters of a shape’s label. | Retrieve the FontSetting, then modify font weight, color, or underline before exporting the workbook.
// AI Prompts: Generate a C# method that accepts a Shape, start index, length, PresetShadowType, and spacing, and returns a configured FontSetting using Aspose.Cells. | Show how to chain additional font properties (size, italic, underline) after obtaining the FontSetting. | Write validation logic that throws ArgumentOutOfRangeException when the character range exceeds the shape's text length.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// A C# helper that retrieves a FontSetting for a specified character range in a Shape, assigns a PresetShadowType, sets character spacing, and returns the configured object for further font styling before saving the workbook.
public static class FontSettingHelper
{
    /// <param name="shape">The shape containing the text.</param>
    /// <param name="startIndex">Zero‑based start index of the characters.</param>
    /// <param name="length">Number of characters to include.</param>
    /// <param name="preset">The preset shadow type to apply.</param>
    /// <param name="spacing">The spacing value to set (positive = wider, negative = tighter).</param>
    /// <returns>A configured FontSetting instance.</returns>
    public static FontSetting CreateFontSettingWithShadowAndSpacing(
        Shape shape,
        int startIndex,
        int length,
        PresetShadowType preset,
        double spacing)
    {
        // Obtain the FontSetting for the requested character range.
        FontSetting fontSetting = shape.Characters(startIndex, length);

        // Configure the shadow effect using the specified preset.
        fontSetting.TextOptions.Shadow.PresetType = preset;

        // Set the character spacing.
        fontSetting.TextOptions.Spacing = spacing;

        return fontSetting;
    }
}

public class Example
{
    public static void Run()
    {
        // Create a workbook and add a shape.
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Shape shape = sheet.Shapes.AddAutoShape(AutoShapeType.Rectangle, 4, 4, 4, 4, 200, 100);
        shape.Text = "Aspose Cells Demo";

        // Apply shadow preset and spacing to the whole text.
        FontSetting fs = FontSettingHelper.CreateFontSettingWithShadowAndSpacing(
            shape,
            0,
            shape.Text.Length,
            PresetShadowType.OffsetBottom,
            2.0);

        // Additional font customizations can be done via fs.Font if needed.
        fs.Font.IsBold = true;
        fs.Font.Color = Color.DarkBlue;

        // Save the workbook.
        workbook.Save("DemoWithShadowAndSpacing.xlsx");
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            Example.Run();
            Console.WriteLine("Workbook created successfully.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
