using System;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsThemeVerification
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define 12 custom theme colors (Background1, Text1, ..., FollowedHyperlink)
            Color[] customColors = new Color[]
            {
                Color.FromArgb(255, 255, 0, 0),     // Background1 - Red
                Color.FromArgb(255, 0, 255, 0),     // Text1 - Green
                Color.FromArgb(255, 0, 0, 255),     // Background2 - Blue
                Color.FromArgb(255, 255, 255, 0),   // Text2 - Yellow
                Color.FromArgb(255, 255, 0, 255),   // Accent1 - Magenta
                Color.FromArgb(255, 0, 255, 255),   // Accent2 - Cyan
                Color.FromArgb(255, 128, 0, 128),   // Accent3 - Purple
                Color.FromArgb(255, 128, 128, 0),   // Accent4 - Olive
                Color.FromArgb(255, 0, 128, 128),   // Accent5 - Teal
                Color.FromArgb(255, 128, 0, 0),     // Accent6 - Maroon
                Color.FromArgb(255, 0, 128, 0),     // Hyperlink - Dark Green
                Color.FromArgb(255, 0, 0, 128)      // FollowedHyperlink - Navy
            };

            // Apply the custom theme to the workbook
            workbook.CustomTheme("MyCustomTheme", customColors);

            // Save the workbook to XLSX
            string filePath = "CustomThemeDemo.xlsx";
            workbook.Save(filePath);

            // Verify that the theme XML inside the XLSX contains the new colors
            VerifyThemeColorsInXml(filePath, customColors);
        }

        static void VerifyThemeColorsInXml(string xlsxPath, Color[] expectedColors)
        {
            // Open the XLSX package as a zip archive
            using (FileStream fs = new FileStream(xlsxPath, FileMode.Open, FileAccess.Read))
            using (ZipArchive archive = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                // Theme XML is stored at xl/theme/theme1.xml
                ZipArchiveEntry themeEntry = archive.GetEntry("xl/theme/theme1.xml");
                if (themeEntry == null)
                {
                    Console.WriteLine("Theme XML not found in the workbook.");
                    return;
                }

                // Read the XML content
                string themeXml;
                using (StreamReader reader = new StreamReader(themeEntry.Open(), Encoding.UTF8))
                {
                    themeXml = reader.ReadToEnd();
                }

                // Check that each expected color appears as a hex value in the XML
                bool allFound = true;
                for (int i = 0; i < expectedColors.Length; i++)
                {
                    // Theme colors are stored as srgbClr elements with hex values (e.g., <a:srgbClr val="FF0000"/>)
                    string hex = expectedColors[i].R.ToString("X2") +
                                 expectedColors[i].G.ToString("X2") +
                                 expectedColors[i].B.ToString("X2");
                    if (!themeXml.Contains($"val=\"{hex}\"", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"Color #{i} (#{hex}) not found in theme XML.");
                        allFound = false;
                    }
                }

                Console.WriteLine(allFound
                    ? "All custom theme colors verified in theme XML."
                    : "Some custom theme colors are missing in theme XML.");
            }
        }
    }
}