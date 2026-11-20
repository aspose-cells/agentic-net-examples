// Title: Modify the Accent1 color in a saved XLSX by editing theme1.xml using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells change accent1 color in exported XLSX | Edit theme1.xml of a workbook saved with Aspose.Cells | How to verify Excel theme XML color change after saving with Aspose.Cells | Update Excel theme accent color programmatically using ZipArchive in .NET | Modify Excel theme XML to set custom color using Aspose.Cells C# example
// Tags: Aspose.Cells edit Excel theme XML | change theme color in XLSX file | save workbook as XLSX then modify theme | use ZipArchive to update theme part in Excel | verify theme XML color after workbook save

using System;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for Theme APIs (if supported)

// The example creates a new Workbook, saves it as ModifiedWorkbook.xlsx, opens the file as a ZipArchive, locates xl/theme/theme1.xml, changes the <a:accent1><a:srgbClr> element's "val" attribute to the hex color 4A90E2, writes the updated XML back into the archive, and confirms the change by comparing the attribute value.
class Program
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook
            Workbook workbook = new Workbook();

            // 2. Save the workbook to XLSX format
            string outputPath = "ModifiedWorkbook.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Ensure the file was created before attempting to modify it
            if (!File.Exists(outputPath))
            {
                Console.WriteLine($"Failed to create the workbook at '{outputPath}'.");
                return;
            }

            // 3. Open the saved XLSX as a zip archive to modify the theme XML
            using (ZipArchive archive = ZipFile.Open(outputPath, ZipArchiveMode.Update))
            {
                // Locate the theme XML part (standard location in XLSX)
                ZipArchiveEntry themeEntry = archive.GetEntry("xl/theme/theme1.xml");
                if (themeEntry == null)
                {
                    Console.WriteLine("Theme XML not found in the workbook.");
                    return;
                }

                // Load the theme XML into an XDocument
                XDocument themeDoc;
                using (Stream themeStream = themeEntry.Open())
                {
                    themeDoc = XDocument.Load(themeStream);
                }

                // Define the drawingML namespace used in the theme XML
                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

                // 4. Retrieve the <a:accent1><a:srgbClr> element
                XElement accent1Element = themeDoc
                    .Root?
                    .Element(a + "themeElements")?
                    .Element(a + "clrScheme")?
                    .Element(a + "accent1")?
                    .Element(a + "srgbClr");

                if (accent1Element == null)
                {
                    Console.WriteLine("Accent1 color element not found in theme XML.");
                    return;
                }

                // 5. Set the desired color value (hex without '#')
                string newColorHex = "4A90E2";
                accent1Element.SetAttributeValue("val", newColorHex);

                // 6. Write the modified XML back into the archive entry
                using (Stream writeStream = themeEntry.Open())
                {
                    // Truncate the existing entry before writing
                    writeStream.SetLength(0);
                    themeDoc.Save(writeStream);
                }

                // 7. Verify that the theme XML now reflects the new color
                string actualColor = accent1Element.Attribute("val")?.Value;
                if (string.Equals(actualColor, newColorHex, StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"Theme verification succeeded: Accent1 color is correctly set to #{newColorHex}.");
                }
                else
                {
                    Console.WriteLine($"Theme verification failed: Expected #{newColorHex}, but found #{actualColor ?? "null"}.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
