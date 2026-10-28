// Title: Add a custom accent color to an Excel workbook theme and verify it after exporting to XLSX with Aspose.Cells for .NET
// Common Searches: how to add a custom accent color to an Excel theme using Aspose.Cells C# | verify that theme1.xml contains a specific hex color after saving a workbook as XLSX | modify Excel theme XML inside the XLSX package with ZipArchive in C# | inject <a:srgbClr> element into theme1.xml using Aspose.Cells | check if a hex color is present in the theme XML of a generated XLSX file
// Tags: add custom accent color to Excel theme with Aspose.Cells | modify theme1.xml in XLSX using ZipArchive C# | verify theme XML after workbook export Aspose.Cells | inject srgbClr element into Excel theme XML | update Excel workbook theme programmatically .NET

using System;
using System.IO;
using System.IO.Compression;
using System.Drawing;
using Aspose.Cells;

// The example creates a workbook, saves it as XLSX, opens the package with ZipArchive, injects an <a:srgbClr> element containing a specified hex color into xl/theme/theme1.xml, overwrites the entry, then reopens the file to confirm the color appears in the theme XML and outputs the verification result.
class ThemeExportVerification
{
    static void Main()
    {
        try
        {
            // Define the new accent color (example: a shade of blue)
            Color newAccentColor = Color.FromArgb(255, 0, 120, 215);
            string hexColor = $"{newAccentColor.R:X2}{newAccentColor.G:X2}{newAccentColor.B:X2}";

            // Create a new workbook and save it as XLSX
            string outputPath = "ModifiedWorkbook.xlsx";
            Workbook workbook = new Workbook();
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Verify that the workbook file exists before attempting to modify it
            if (!File.Exists(outputPath))
            {
                Console.WriteLine("Failed to create the workbook file.");
                return;
            }

            // Open the XLSX package and inject the new color into the theme XML
            using (FileStream fs = new FileStream(outputPath, FileMode.Open, FileAccess.ReadWrite))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Update))
            {
                ZipArchiveEntry themeEntry = zip.GetEntry("xl/theme/theme1.xml");
                if (themeEntry != null)
                {
                    string themeXml;
                    using (StreamReader reader = new StreamReader(themeEntry.Open()))
                    {
                        themeXml = reader.ReadToEnd();
                    }

                    // Inject a new color element containing the desired hex value
                    // The insertion point is before the closing </a:theme> tag
                    string injection = $"<a:srgbClr val=\"{hexColor}\"/>";
                    int insertPos = themeXml.LastIndexOf("</a:theme>", StringComparison.OrdinalIgnoreCase);
                    if (insertPos >= 0)
                    {
                        themeXml = themeXml.Insert(insertPos, injection);
                    }

                    // Overwrite the existing theme entry with the modified XML
                    using (StreamWriter writer = new StreamWriter(themeEntry.Open()))
                    {
                        writer.Write(themeXml);
                    }
                }
                else
                {
                    Console.WriteLine("Theme XML not found in the workbook package.");
                }
            }

            // Re-open the modified workbook to verify the presence of the new color in the theme XML
            bool colorFound = false;
            using (FileStream fs = new FileStream(outputPath, FileMode.Open, FileAccess.Read))
            using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
            {
                ZipArchiveEntry themeEntry = zip.GetEntry("xl/theme/theme1.xml");
                if (themeEntry != null)
                {
                    using (StreamReader reader = new StreamReader(themeEntry.Open()))
                    {
                        string themeXml = reader.ReadToEnd();
                        if (themeXml.IndexOf(hexColor, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            colorFound = true;
                        }
                    }
                }
            }

            // Output verification result
            Console.WriteLine(colorFound
                ? "Verification succeeded: new theme color is present in the theme XML."
                : "Verification failed: new theme color not found in the theme XML.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
