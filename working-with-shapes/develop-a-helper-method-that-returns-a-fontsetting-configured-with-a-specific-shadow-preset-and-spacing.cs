using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

public static class FontSettingHelper
{
    /// <summary>
    /// Creates a FontSetting for a shape's text, applies the specified shadow preset and character spacing,
    /// and returns the configured FontSetting.
    /// </summary>
    /// <param name="preset">The preset shadow type to apply.</param>
    /// <param name="spacing">The spacing value between characters.</param>
    /// <returns>A FontSetting object with the desired shadow and spacing settings.</returns>
    public static FontSetting GetFontSettingWithShadowAndSpacing(PresetShadowType preset, double spacing)
    {
        try
        {
            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle auto shape to the worksheet.
            Shape shape = sheet.Shapes.AddAutoShape(AutoShapeType.Rectangle, 1, 1, 0, 0, 200, 100);

            // Set sample text for the shape.
            shape.Text = "Sample Text";

            // Obtain a FontSetting that covers the entire text of the shape.
            FontSetting fontSetting = shape.Characters(0, shape.Text.Length);

            // Configure the shadow effect using the provided preset.
            fontSetting.TextOptions.Shadow.PresetType = preset;

            // Set the character spacing.
            fontSetting.TextOptions.Spacing = spacing;

            return fontSetting;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in GetFontSettingWithShadowAndSpacing: {ex.Message}");
            throw;
        }
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Example usage of the helper method.
            var fontSetting = FontSettingHelper.GetFontSettingWithShadowAndSpacing(
                PresetShadowType.OffsetDiagonalBottomRight, 2.0);
            Console.WriteLine("FontSetting configured successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}