using System;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;

class CustomThemeExportAndVerify
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Define 12 custom theme colors (Background1, Text1, ..., FollowedHyperlink)
        Color[] customColors = new Color[]
        {
            Color.FromArgb(255, 255, 0, 0),     // Red
            Color.FromArgb(255, 0, 255, 0),     // Green
            Color.FromArgb(255, 0, 0, 255),     // Blue
            Color.FromArgb(255, 255, 255, 0),   // Yellow
            Color.FromArgb(255, 255, 0, 255),   // Magenta
            Color.FromArgb(255, 0, 255, 255),   // Cyan
            Color.FromArgb(255, 128, 0, 128),   // Purple
            Color.FromArgb(255, 128, 128, 0),   // Olive
            Color.FromArgb(255, 0, 128, 128),   // Teal
            Color.FromArgb(255, 128, 0, 0),     // Maroon
            Color.FromArgb(255, 0, 128, 0),     // Dark Green
            Color.FromArgb(255, 0, 0, 128)      // Navy
        };

        // Apply the custom theme to the workbook
        workbook.CustomTheme("MyCustomTheme", customColors);

        // Save the workbook as XLSX
        string filePath = "CustomThemeDemo.xlsx";
        workbook.Save(filePath);

        // Verify that the theme XML inside the saved XLSX contains the new colors
        using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
        {
            // Theme XML is stored at xl/theme/theme1.xml
            ZipArchiveEntry themeEntry = zip.GetEntry("xl/theme/theme1.xml");
            if (themeEntry == null)
            {
                Console.WriteLine("Theme XML not found in the workbook.");
                return;
            }

            string themeXml;
            using (StreamReader reader = new StreamReader(themeEntry.Open()))
            {
                themeXml = reader.ReadToEnd();
            }

            // Check for the presence of the first custom color (Red) in the XML.
            // Aspose stores colors as six‑digit hex without the leading '#'.
            string firstColorHex = ColorTranslator.ToHtml(customColors[0]).TrimStart('#').ToUpperInvariant();

            bool containsFirstColor = themeXml.IndexOf(firstColorHex, StringComparison.OrdinalIgnoreCase) >= 0;
            Console.WriteLine($"Theme XML contains first custom color (#{firstColorHex}): {containsFirstColor}");

            // Optionally, verify all custom colors are present
            bool allColorsPresent = true;
            for (int i = 0; i < customColors.Length; i++)
            {
                string hex = ColorTranslator.ToHtml(customColors[i]).TrimStart('#').ToUpperInvariant();
                if (themeXml.IndexOf(hex, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    allColorsPresent = false;
                    Console.WriteLine($"Missing color #{hex} in theme XML (index {i}).");
                }
            }
            Console.WriteLine($"All custom colors present in theme XML: {allColorsPresent}");
        }
    }
}