// Title: C# helper to create Aspose.Cells FontSetting with preset shadow and character spacing
// Description: A concise C# method that builds a FontSetting for a specific character range, applies a chosen PresetShadowType, sets custom character spacing, and returns the configured object for further styling in Aspose.Cells workbooks.
// Keywords: Aspose.Cells | FontSetting | preset shadow | character spacing | C# | Aspose.Cells TextOptions | Shadow.PresetType | worksheet text formatting | Aspose.Cells example | GitHub | dotnet
// Common Searches: Aspose.Cells set preset shadow on FontSetting | C# change character spacing in Excel cell using Aspose | How to use FontSetting TextOptions in Aspose.Cells | Apply shadow effect to part of cell text Aspose.Cells | Sample code FontSetting shadow and spacing
// Developer Intent: Generate a FontSetting that applies a selected shadow preset and custom spacing to a defined text range in an Aspose.Cells worksheet.
// Use Cases: Highlight the first few characters of a header with a bottom‑offset shadow and wider spacing for visual emphasis. | Create a reusable helper that formats column titles across multiple sheets, each with distinct shadow presets and tighter spacing. | Produce a styled report where section titles receive a soft shadow and expanded spacing while body text remains default.
// AI Prompts: Write a C# method that returns a FontSetting with a given PresetShadowType and spacing for a specified character range in Aspose.Cells. | Show how to call FontSettingHelper.GetFontSettingWithShadowAndSpacing to style text in a workbook and then save the file. | Explain how to extend the helper to also set underline style and font color while preserving shadow and spacing settings.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

// A concise C# method that builds a FontSetting for a specific character range, applies a chosen PresetShadowType, sets custom character spacing, and returns the configured object for further styling in Aspose.Cells workbooks.
public static class FontSettingHelper
{
    /// <param name="sheets">The collection of worksheets the FontSetting belongs to.</param>
    /// <param name="startIndex">Zero‑based start index of the character range.</param>
    /// <param name="length">Number of characters in the range.</param>
    /// <param name="preset">The preset shadow type to apply.</param>
    /// <param name="spacing">The spacing value to set (positive = wider, negative = tighter).</param>
    /// <returns>A configured FontSetting instance.</returns>
    public static FontSetting GetFontSettingWithShadowAndSpacing(
        WorksheetCollection sheets,
        int startIndex,
        int length,
        PresetShadowType preset,
        double spacing)
    {
        // Create the FontSetting for the specified character range.
        FontSetting fontSetting = new FontSetting(startIndex, length, sheets);

        // Access the TextOptions associated with this FontSetting.
        TextOptions textOptions = fontSetting.TextOptions;

        // Configure the shadow effect using the provided preset.
        textOptions.Shadow.PresetType = preset;

        // Set the character spacing.
        textOptions.Spacing = spacing;

        // Return the fully configured FontSetting.
        return fontSetting;
    }
}

public class Program
{
    public static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet.
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Set a sample value in cell A1.
            ws.Cells["A1"].PutValue("Hello World");

            // Apply font settings to the first 5 characters of the cell.
            FontSetting fs = FontSettingHelper.GetFontSettingWithShadowAndSpacing(
                wb.Worksheets, 0, 5, PresetShadowType.OffsetBottom, 2.0);

            // Additional font styling.
            fs.Font.IsBold = true;

            // Define output file path.
            string outputPath = "output.xlsx";

            // Ensure the output directory exists.
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook.
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
