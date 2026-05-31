using System;
using System.Drawing;
using Aspose.Cells;

namespace CustomThemeTemplate
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (macro‑free by default)
            Workbook workbook = new Workbook();

            // Define 12 theme colors (Background1, Text1, Background2, Text2, Accent1‑Accent6, Hyperlink, FollowedHyperlink)
            Color[] themeColors = new Color[]
            {
                Color.FromArgb(255, 255, 255), // Background1 - White
                Color.FromArgb(0,   0,   0),   // Text1       - Black
                Color.FromArgb(242, 242, 242), // Background2 - Light Gray
                Color.FromArgb(64,  64,  64),   // Text2       - Dark Gray
                Color.FromArgb(0,   112, 192), // Accent1     - Blue
                Color.FromArgb(255, 192, 0),   // Accent2     - Orange
                Color.FromArgb(112, 173, 71),  // Accent3     - Green
                Color.FromArgb(255, 0,   0),   // Accent4     - Red
                Color.FromArgb(255, 0, 255),   // Accent5     - Magenta
                Color.FromArgb(0,   255, 255), // Accent6     - Cyan
                Color.FromArgb(0,   0, 255),   // Hyperlink   - Blue
                Color.FromArgb(128, 0,   128)  // FollowedHyperlink - Purple
            };

            // Apply the custom theme to the workbook
            workbook.CustomTheme("MyTemplateTheme", themeColors);

            // Optional: demonstrate using a theme color in a cell style
            Worksheet sheet = workbook.Worksheets[0];
            Cell demoCell = sheet.Cells["A1"];
            demoCell.PutValue("Themed Text");

            Style themedStyle = workbook.CreateStyle();
            // Use Accent1 from the custom theme
            themedStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
            themedStyle.Font.Size = 12;
            demoCell.SetStyle(themedStyle);

            // Ensure the workbook is macro‑free (no effect on a newly created workbook but kept for safety)
            workbook.RemoveMacro();

            // Save the template workbook
            string outputPath = "TemplateWithCustomTheme.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"Template workbook saved to '{outputPath}' with custom theme.");
        }
    }
}