using Aspose.Cells;
using System.Drawing;

class CustomThemeDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Prepare an array of 12 theme colors.
        // Indexes 4 and 5 correspond to Accent1 and Accent2.
        Color[] customColors = new Color[12];
        customColors[0] = Color.White;                     // Background1
        customColors[1] = Color.Black;                     // Text1
        customColors[2] = Color.White;                     // Background2
        customColors[3] = Color.Black;                     // Text2
        customColors[4] = Color.FromArgb(255, 128, 0);      // Accent1 – custom orange
        customColors[5] = Color.FromArgb(0, 128, 255);      // Accent2 – custom light blue
        customColors[6] = Color.Gray;                      // Accent3
        customColors[7] = Color.Green;                     // Accent4
        customColors[8] = Color.Purple;                    // Accent5
        customColors[9] = Color.Brown;                     // Accent6
        customColors[10] = Color.Blue;                     // Hyperlink
        customColors[11] = Color.Red;                      // Followed Hyperlink

        // Apply the custom theme to the workbook
        workbook.CustomTheme("MyCustomTheme", customColors);

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Demonstrate Accent1 as a font theme color
        Cell fontCell = sheet.Cells["A1"];
        fontCell.PutValue("Accent1 Font");
        Style fontStyle = workbook.CreateStyle();
        fontStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
        fontCell.SetStyle(fontStyle);

        // Demonstrate Accent2 as a fill (foreground) theme color
        Cell fillCell = sheet.Cells["A2"];
        fillCell.PutValue("Accent2 Fill");
        Style fillStyle = workbook.CreateStyle();
        fillStyle.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent2, 0.0);
        fillStyle.Pattern = BackgroundType.Solid;
        fillCell.SetStyle(fillStyle);

        // Save the workbook
        workbook.Save("CustomThemeAccentDemo.xlsx");
    }
}