using System;
using System.Drawing;
using Aspose.Cells;

namespace UpdateAccent4Theme
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Prompt user for a hex color (e.g., #FF5733)
            Console.Write("Enter a hex color for Accent4 (e.g., #FF5733): ");
            string hexInput = Console.ReadLine();

            // Convert the hex string to a System.Drawing.Color
            Color accent4Color;
            try
            {
                accent4Color = ColorTranslator.FromHtml(hexInput);
            }
            catch
            {
                Console.WriteLine("Invalid color format. Using default Red.");
                accent4Color = Color.Red;
            }

            // Update the workbook's Accent4 theme color (feature rule: SetThemeColor)
            workbook.SetThemeColor(ThemeColorType.Accent4, accent4Color);

            // Refresh all cell styles that depend on Accent4
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all used cells in the worksheet
                foreach (Cell cell in sheet.Cells)
                {
                    Style style = cell.GetStyle();
                    bool needsUpdate = false;

                    // Font theme color
                    if (style.Font.ThemeColor != null &&
                        style.Font.ThemeColor.ColorType == ThemeColorType.Accent4)
                    {
                        // Reassign to trigger refresh (preserve existing tint)
                        double tint = style.Font.ThemeColor.Tint;
                        style.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent4, tint);
                        needsUpdate = true;
                    }

                    // Foreground theme color
                    if (style.ForegroundThemeColor != null &&
                        style.ForegroundThemeColor.ColorType == ThemeColorType.Accent4)
                    {
                        double tint = style.ForegroundThemeColor.Tint;
                        style.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent4, tint);
                        needsUpdate = true;
                    }

                    // Background theme color
                    if (style.BackgroundThemeColor != null &&
                        style.BackgroundThemeColor.ColorType == ThemeColorType.Accent4)
                    {
                        double tint = style.BackgroundThemeColor.Tint;
                        style.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent4, tint);
                        needsUpdate = true;
                    }

                    // Apply the refreshed style back to the cell if any changes were made
                    if (needsUpdate)
                    {
                        cell.SetStyle(style);
                    }
                }
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save("UpdatedAccent4Theme.xlsx");
            Console.WriteLine("Workbook saved as UpdatedAccent4Theme.xlsx");
        }
    }
}