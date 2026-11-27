// Title: C# helper method to create a FontSetting for partial text formatting in Aspose.Cells (bold, color, placeholder for shadow and spacing)
// Common Searches: how to format a substring of a cell using FontSetting in Aspose.Cells C# | Aspose.Cells ReplaceOptions with custom FontSetting example | C# create FontSetting for specific character range in a workbook | apply bold and dark blue color to selected characters with Aspose.Cells | use FontSetting to prepare for shadow or spacing settings in Aspose.Cells
// Tags: Aspose.Cells FontSetting partial text formatting | C# FontSetting ReplaceOptions usage | bold color character range Aspose.Cells | placeholder shadow type FontSetting Aspose.Cells | character spacing preparation FontSetting Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

namespace FontSettingHelperDemo
{
    // The example defines a GetFontSettingWithShadowAndSpacing helper that builds a FontSetting for a given start index and length, applies bold and dark blue styling, and returns it for use with ReplaceOptions to replace text in a worksheet. It notes that current Aspose.Cells versions do not expose direct shadow or spacing properties, so those are placeholders for future API support.
    class Program
    {
        static FontSetting GetFontSettingWithShadowAndSpacing(Workbook workbook, int startIndex, int length, PresetShadowType shadowType, double spacing)
        {
            // FontSetting constructor requires start index, length and the worksheets collection.
            FontSetting fontSetting = new FontSetting(startIndex, length, workbook.Worksheets);

            // Basic font attributes for visual effect.
            fontSetting.Font.IsBold = true;
            fontSetting.Font.Color = Color.DarkBlue;

            // Note: Aspose.Cells' FontSetting does not currently expose shadow or spacing properties.
            // Future versions may allow setting these via other APIs.

            return fontSetting;
        }

        static void Main(string[] args)
        {
            const string outputPath = "FontSettingWithShadowAndSpacing.xlsx";

            try
            {
                // 1. Create a new workbook and get the first worksheet.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Put sample text into a cell.
                sheet.Cells["A1"].PutValue("This is a placeholder text that will be replaced.");

                // 3. Prepare ReplaceOptions and attach the FontSetting.
                ReplaceOptions replaceOptions = new ReplaceOptions
                {
                    MatchEntireCellContents = false
                };

                // The word "placeholder" starts at index 10 and has length 11.
                FontSetting formattedSetting = GetFontSettingWithShadowAndSpacing(
                    workbook,
                    startIndex: 10,
                    length: 11,
                    shadowType: PresetShadowType.OffsetDiagonalBottomRight,
                    spacing: 2.0);

                replaceOptions.FontSettings = new FontSetting[] { formattedSetting };

                // 4. Perform the replacement.
                workbook.Replace("placeholder", "Formatted", replaceOptions);

                // 5. Save the workbook.
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
