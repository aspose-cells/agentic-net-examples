// Title: Read an Excel workbook’s theme font scheme and output the primary (major) and secondary (minor) font families with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file using Aspose.Cells, accesses Workbook.Theme.FontScheme, and prints the LatinTypeface of the MajorFont and MinorFont. | Extend the sample to also retrieve the EastAsianTypeface for both major and minor fonts, and return the workbook's default style font when the Theme object is absent. | Create a reusable C# method that returns a tuple (primaryFont, secondaryFont) from a given Excel file, employing reflection to handle environments where the Theme class is unavailable.
// Common Searches: aspocells get theme major font family c# | how to read Excel theme font scheme using Aspose.Cells .NET | use default workbook style font if theme information is missing Aspose.Cells | c# reflection access Workbook.Theme property Aspose.Cells example | retrieve primary and secondary font families from Excel file with Aspose.Cells
// Tags: Aspose.Cells read theme font scheme | C# extract major and minor font families | fallback to workbook default style font | reflection access Workbook Theme property | retrieve LatinTypeface from Excel theme

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, uses reflection to safely obtain Workbook.Theme.FontScheme, extracts the LatinTypeface of the major (primary) and minor (secondary) fonts, falls back to the workbook's default style font when the theme is unavailable, and prints the font family names.
class Program
{
    static void Main()
    {
        string filePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(filePath);

            // Attempt to retrieve theme font information if supported
            string primaryFont = "N/A";
            string secondaryFont = "N/A";

            // The Theme class is available in newer versions of Aspose.Cells.
            // Use reflection to safely access it without causing compile‑time errors
            var themeProperty = typeof(Workbook).GetProperty("Theme");
            if (themeProperty != null)
            {
                var theme = themeProperty.GetValue(workbook);
                if (theme != null)
                {
                    var fontSchemeProp = theme.GetType().GetProperty("FontScheme");
                    var fontScheme = fontSchemeProp?.GetValue(theme);
                    if (fontScheme != null)
                    {
                        var majorFontProp = fontScheme.GetType().GetProperty("MajorFont");
                        var minorFontProp = fontScheme.GetType().GetProperty("MinorFont");

                        var majorFont = majorFontProp?.GetValue(fontScheme);
                        var minorFont = minorFontProp?.GetValue(fontScheme);

                        var latinTypefaceProp = typeof(object).GetProperty("LatinTypeface"); // placeholder; will be resolved via dynamic

                        if (majorFont != null)
                        {
                            var latinProp = majorFont.GetType().GetProperty("LatinTypeface");
                            primaryFont = latinProp?.GetValue(majorFont)?.ToString() ?? primaryFont;
                        }

                        if (minorFont != null)
                        {
                            var latinProp = minorFont.GetType().GetProperty("LatinTypeface");
                            secondaryFont = latinProp?.GetValue(minorFont)?.ToString() ?? secondaryFont;
                        }
                    }
                }
            }

            // Fallback to default style if theme information is unavailable
            if (primaryFont == "N/A")
            {
                primaryFont = workbook.DefaultStyle.Font.Name;
            }

            // Output the font families
            Console.WriteLine("Primary Font Family: " + primaryFont);
            Console.WriteLine("Secondary Font Family: " + secondaryFont);
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
