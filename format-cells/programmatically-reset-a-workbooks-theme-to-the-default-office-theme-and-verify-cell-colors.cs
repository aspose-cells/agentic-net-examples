using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeResetDemo
{
    class Program
    {
        static void Main()
        {
            // ------------------------------------------------------------
            // 1. Create a new workbook (default Office theme is applied)
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // 2. Apply a custom theme color (e.g., change Accent1 to Red)
            // ------------------------------------------------------------
            workbook.SetThemeColor(ThemeColorType.Accent1, Color.Red);

            // Create a style that uses the Accent1 theme color
            Style customStyle = workbook.CreateStyle();
            customStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
            customStyle.Font.Size = 12;
            customStyle.Font.IsBold = true;
            customStyle.Pattern = BackgroundType.Solid;
            customStyle.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
            sheet.Cells["A1"].PutValue("Custom Theme Cell");
            sheet.Cells["A1"].SetStyle(customStyle);

            // Save the workbook with the custom theme
            workbook.Save("CustomTheme.xlsx");

            // ------------------------------------------------------------
            // 3. Reset the workbook's theme to the default Office theme
            //    by copying the theme from a freshly created workbook.
            // ------------------------------------------------------------
            Workbook defaultThemeWorkbook = new Workbook(); // contains the default theme
            workbook.CopyTheme(defaultThemeWorkbook);      // reset theme

            // ------------------------------------------------------------
            // 4. Verify that the theme color has been restored to the default
            //    (compare Accent1 color before and after reset)
            // ------------------------------------------------------------
            Color defaultAccent1 = defaultThemeWorkbook.GetThemeColor(ThemeColorType.Accent1);
            Color currentAccent1 = workbook.GetThemeColor(ThemeColorType.Accent1);

            Console.WriteLine($"Default Accent1 Color:  A={defaultAccent1.A}, R={defaultAccent1.R}, G={defaultAccent1.G}, B={defaultAccent1.B}");
            Console.WriteLine($"Current Accent1 Color:  A={currentAccent1.A}, R={currentAccent1.R}, G={currentAccent1.G}, B={currentAccent1.B}");

            bool isReset = defaultAccent1.ToArgb() == currentAccent1.ToArgb();
            Console.WriteLine($"Theme reset successful: {isReset}");

            // ------------------------------------------------------------
            // 5. Save the workbook after resetting the theme
            // ------------------------------------------------------------
            workbook.Save("ResetTheme.xlsx");
        }
    }
}