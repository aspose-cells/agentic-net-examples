using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (macro‑free by default)
        Workbook workbook = new Workbook();

        // Ensure any macros are removed (extra safety)
        workbook.RemoveMacro();

        // Define 12 custom theme colors (Background1, Text1, Background2, Text2, Accent1‑Accent6, Hyperlink, FollowedHyperlink)
        Color[] customColors = new Color[]
        {
            Color.FromArgb(255, 255, 255), // Background1 – white
            Color.FromArgb(0,   0,   0),   // Text1 – black
            Color.FromArgb(242, 242, 242), // Background2 – light gray
            Color.FromArgb(64,  64,  64),   // Text2 – dark gray
            Color.FromArgb(0,   112, 192), // Accent1 – blue
            Color.FromArgb(255, 192, 0),   // Accent2 – orange
            Color.FromArgb(112, 173, 71),  // Accent3 – green
            Color.FromArgb(255, 0,   0),   // Accent4 – red
            Color.FromArgb(255, 0, 255),   // Accent5 – magenta
            Color.FromArgb(0,   255, 255), // Accent6 – cyan
            Color.FromArgb(0,   0, 255),   // Hyperlink – blue
            Color.FromArgb(128, 0, 128)    // Followed Hyperlink – purple
        };

        // Apply the custom theme to the workbook
        workbook.CustomTheme("MyCustomTheme", customColors);

        // (Optional) Demonstrate usage of a theme color in a cell style
        Worksheet sheet = workbook.Worksheets[0];
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Themed Cell");

        Style style = workbook.CreateStyle();
        // Use Accent1 from the custom theme
        style.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
        style.Font.Size = 12;
        cell.SetStyle(style);

        // Save the macro‑free template workbook
        workbook.Save("TemplateWithCustomTheme.xlsx");
    }
}