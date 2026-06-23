using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

namespace ThemePaletteUtility
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourceFile = "ThemeSource.xlsx";
                const string outputFile = "GeneratedPalette.xlsx";

                // Verify that the source Excel file exists
                if (!File.Exists(sourceFile))
                {
                    Console.WriteLine($"Source file not found: {sourceFile}");
                    return;
                }

                // Load the workbook that contains the theme
                Workbook sourceWorkbook = new Workbook(sourceFile);

                // Retrieve the six accent colors from the workbook's theme
                Color[] accentColors = new Color[6];
                accentColors[0] = sourceWorkbook.GetThemeColor(ThemeColorType.Accent1);
                accentColors[1] = sourceWorkbook.GetThemeColor(ThemeColorType.Accent2);
                accentColors[2] = sourceWorkbook.GetThemeColor(ThemeColorType.Accent3);
                accentColors[3] = sourceWorkbook.GetThemeColor(ThemeColorType.Accent4);
                accentColors[4] = sourceWorkbook.GetThemeColor(ThemeColorType.Accent5);
                accentColors[5] = sourceWorkbook.GetThemeColor(ThemeColorType.Accent6);

                // Create a new workbook to demonstrate the palette
                Workbook destWorkbook = new Workbook();
                Worksheet sheet = destWorkbook.Worksheets[0];
                sheet.Name = "Accent Palette";

                // Populate cells with colored rectangles and labels
                for (int i = 0; i < accentColors.Length; i++)
                {
                    // Set cell value as label
                    Cell cell = sheet.Cells[i, 0];
                    cell.PutValue($"Accent{i + 1}");

                    // Apply solid fill with the accent color
                    Style style = cell.GetStyle();
                    style.ForegroundColor = accentColors[i];
                    style.Pattern = BackgroundType.Solid;
                    cell.SetStyle(style);
                }

                // Auto-fit the column for better visibility
                sheet.AutoFitColumn(0);

                // Save the generated workbook
                destWorkbook.Save(outputFile, SaveFormat.Xlsx);
                Console.WriteLine($"Palette workbook saved as {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}