using Aspose.Cells;
using System.Drawing;

class CustomThemeDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Prepare an array of 12 theme colors (indices 0‑11)
        Color[] customColors = new Color[12];

        // Default theme colors (can be any valid colors)
        customColors[0] = Color.White;                     // Background1
        customColors[1] = Color.Black;                     // Text1
        customColors[2] = Color.LightGray;                 // Background2
        customColors[3] = Color.DarkGray;                  // Text2

        // Custom Accent1 and Accent2 colors (new RGB values)
        customColors[4] = Color.FromArgb(255, 128, 0);      // Accent1 – orange
        customColors[5] = Color.FromArgb(0, 128, 255);      // Accent2 – light blue

        // Remaining theme colors (placeholders)
        customColors[6] = Color.Green;                     // Accent3
        customColors[7] = Color.Purple;                    // Accent4
        customColors[8] = Color.Brown;                     // Accent5
        customColors[9] = Color.Magenta;                   // Accent6
        customColors[10] = Color.Blue;                     // Hyperlink
        customColors[11] = Color.Purple;                   // Followed Hyperlink

        // Apply the custom theme to the workbook
        workbook.CustomTheme("MyCustomTheme", customColors);

        // Demonstrate the new Accent1 color in a cell
        Worksheet sheet = workbook.Worksheets[0];
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Accent1 Theme Text");

        // Use Accent1 as the font theme color
        Style style = workbook.CreateStyle();
        style.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
        cell.SetStyle(style);

        // Save the workbook
        workbook.Save("CustomThemeAccentDemo.xlsx");
    }
}