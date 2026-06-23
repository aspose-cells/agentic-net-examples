using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;

class ThemeChangeLogger
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Define cells to monitor and assign them initial theme colors
        List<string> cellAddresses = new List<string> { "A1", "A2", "A3", "A4" };
        ThemeColorType[] themeTypes = new ThemeColorType[]
        {
            ThemeColorType.Accent1,
            ThemeColorType.Accent2,
            ThemeColorType.Accent3,
            ThemeColorType.Accent4
        };

        for (int i = 0; i < cellAddresses.Count; i++)
        {
            Cell cell = worksheet.Cells[cellAddresses[i]];
            cell.PutValue($"Theme cell {i + 1}");

            Style style = workbook.CreateStyle();
            style.Font.ThemeColor = new ThemeColor(themeTypes[i], 0.0); // initial tint 0
            cell.SetStyle(style);
        }

        // Capture the original ThemeColor of each watched cell
        Dictionary<string, ThemeColor> oldThemeColors = new Dictionary<string, ThemeColor>();
        foreach (string addr in cellAddresses)
        {
            Style style = worksheet.Cells[addr].GetStyle();
            oldThemeColors[addr] = style.Font.ThemeColor;
        }

        // Prepare a custom theme (12 colors required)
        Color[] customColors = new Color[12];
        customColors[0] = Color.White;          // Background1
        customColors[1] = Color.Black;          // Text1
        customColors[2] = Color.LightGray;      // Background2
        customColors[3] = Color.DarkGray;       // Text2
        customColors[4] = Color.Orange;         // Accent1
        customColors[5] = Color.Purple;         // Accent2
        customColors[6] = Color.Teal;           // Accent3
        customColors[7] = Color.Brown;          // Accent4
        customColors[8] = Color.Pink;           // Accent5
        customColors[9] = Color.Yellow;         // Accent6
        customColors[10] = Color.Blue;          // Hyperlink
        customColors[11] = Color.Red;           // Followed Hyperlink

        // Apply the custom theme (lifecycle rule: modify existing workbook)
        workbook.CustomTheme("MyCustomTheme", customColors);

        // Capture the ThemeColor after applying the theme
        Dictionary<string, ThemeColor> newThemeColors = new Dictionary<string, ThemeColor>();
        foreach (string addr in cellAddresses)
        {
            Style style = worksheet.Cells[addr].GetStyle();
            newThemeColors[addr] = style.Font.ThemeColor;
        }

        // Log cells whose ThemeColor did not change
        Console.WriteLine("Cells that failed to update after theme change:");
        bool anyFailure = false;
        foreach (string addr in cellAddresses)
        {
            ThemeColor oldTc = oldThemeColors[addr];
            ThemeColor newTc = newThemeColors[addr];

            // Compare both ColorType and Tint (allow tiny floating‑point differences)
            if (oldTc.ColorType == newTc.ColorType && Math.Abs(oldTc.Tint - newTc.Tint) < 0.0001)
            {
                Console.WriteLine($"{addr} - ThemeColor unchanged (Type={oldTc.ColorType}, Tint={oldTc.Tint})");
                anyFailure = true;
            }
        }

        if (!anyFailure)
        {
            Console.WriteLine("All watched cells updated correctly.");
        }

        // Save the workbook (lifecycle rule: save)
        workbook.Save("ThemeChangeLogDemo.xlsx");
    }
}