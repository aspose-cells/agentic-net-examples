// Title: Export a Workbook with a Custom Theme to XLSX and Validate Theme XML Colors using Aspose.Cells for .NET (C#)
// Description: Creates a new Workbook, defines a 12‑color custom theme, applies it via workbook.CustomTheme, saves the file as XLSX, then extracts xl/theme/theme1.xml from the package and uses regular expressions to confirm each RGB value appears as an a:srgbClr entry.
// Keywords: Aspose.Cells custom theme | C# export XLSX | verify theme XML | theme1.xml color check | apply brand palette Aspose.Cells | read XLSX zip archive | a:srgbClr validation | .NET spreadsheet theming
// Common Searches: how to apply a custom theme with Aspose.Cells and save as XLSX | C# read theme1.xml from generated Excel file | verify custom theme colors in exported XLSX | Aspose.Cells export workbook with custom palette | check theme XML colors programmatically
// Developer Intent: Generate an XLSX file that uses a user‑defined 12‑color theme and programmatically ensure the theme part contains the exact RGB codes.
// Use Cases: Implement brand‑consistent spreadsheets by embedding a corporate color scheme as a custom theme. | Automated regression testing to guarantee that theme export retains precise color values. | Create sample documents for documentation or training that require a predefined visual style.
// AI Prompts: Write C# code with Aspose.Cells to define a 12‑color custom theme, apply it, save as XLSX, and verify the theme XML colors. | Show how to open a saved XLSX, locate xl/theme/theme1.xml, and use regex to confirm specific a:srgbClr values. | Explain the role of workbook.CustomTheme and how to programmatically validate that the theme part was written correctly.

using System;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using Aspose.Cells;

// Creates a new Workbook, defines a 12‑color custom theme, applies it via workbook.CustomTheme, saves the file as XLSX, then extracts xl/theme/theme1.xml from the package and uses regular expressions to confirm each RGB value appears as an a:srgbClr entry.
class Program
{
    static void Main()
    {
        // Create a new workbook and add a sample cell
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Theme Test");

        // Define 12 custom theme colors (order matches ThemeColorType enumeration)
        Color[] customColors = new Color[]
        {
            Color.FromArgb(255, 255, 0, 0),   // Background1 (Red)
            Color.FromArgb(255, 0, 255, 0),   // Text1 (Green)
            Color.FromArgb(255, 0, 0, 255),   // Background2 (Blue)
            Color.FromArgb(255, 255, 255, 0), // Text2 (Yellow)
            Color.FromArgb(255, 255, 0, 255), // Accent1 (Magenta)
            Color.FromArgb(255, 0, 255, 255), // Accent2 (Cyan)
            Color.FromArgb(255, 128, 0, 128), // Accent3 (Purple)
            Color.FromArgb(255, 128, 128, 0), // Accent4 (Olive)
            Color.FromArgb(255, 0, 128, 128), // Accent5 (Teal)
            Color.FromArgb(255, 128, 0, 0),   // Accent6 (Maroon)
            Color.FromArgb(255, 0, 128, 0),   // Hyperlink (DarkGreen)
            Color.FromArgb(255, 0, 0, 128)    // FollowedHyperlink (Navy)
        };

        // Apply the custom theme to the workbook
        workbook.CustomTheme("MyCustomTheme", customColors);

        // Save the workbook as XLSX
        string filePath = "CustomThemeDemo.xlsx";
        workbook.Save(filePath);

        // Verify that the theme XML inside the saved XLSX contains the new colors
        bool verificationPassed = VerifyThemeColors(filePath, customColors);
        Console.WriteLine("Theme verification " + (verificationPassed ? "passed" : "failed"));
    }

    // Reads the theme XML part from the XLSX package and checks for each expected color
    static bool VerifyThemeColors(string xlsxPath, Color[] expectedColors)
    {
        using (FileStream fs = new FileStream(xlsxPath, FileMode.Open, FileAccess.Read))
        using (ZipArchive archive = new ZipArchive(fs, ZipArchiveMode.Read))
        {
            // Theme XML is typically stored at xl/theme/theme1.xml
            ZipArchiveEntry themeEntry = archive.GetEntry("xl/theme/theme1.xml");
            if (themeEntry == null)
            {
                Console.WriteLine("Theme XML part not found.");
                return false;
            }

            string themeXml;
            using (StreamReader reader = new StreamReader(themeEntry.Open()))
            {
                themeXml = reader.ReadToEnd();
            }

            // Each color is stored as a:srgbClr val="RRGGBB"
            foreach (Color color in expectedColors)
            {
                string hex = color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
                string pattern = $"val=\"{hex}\"";
                if (!Regex.IsMatch(themeXml, pattern, RegexOptions.IgnoreCase))
                {
                    Console.WriteLine($"Color {hex} not found in theme XML.");
                    return false;
                }
            }
        }
        return true;
    }
}
