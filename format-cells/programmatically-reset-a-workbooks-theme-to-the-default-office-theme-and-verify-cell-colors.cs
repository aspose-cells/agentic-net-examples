using System;
using System.Drawing;
using Aspose.Cells;

class ResetThemeDemo
{
    static void Main()
    {
        // Create a workbook and apply a custom theme color (Accent1 = Red)
        Workbook workbook = new Workbook();
        workbook.SetThemeColor(ThemeColorType.Accent1, Color.Red);
        workbook.Save("CustomTheme.xlsx"); // optional: save before reset

        // Create a workbook that has the default Office theme
        Workbook defaultWorkbook = new Workbook(); // default theme is applied automatically

        // Reset the original workbook's theme by copying the default theme
        workbook.CopyTheme(defaultWorkbook);

        // Verify that the theme colors have been restored to the defaults
        Color currentAccent1 = workbook.GetThemeColor(ThemeColorType.Accent1);
        Color defaultAccent1 = defaultWorkbook.GetThemeColor(ThemeColorType.Accent1);

        Console.WriteLine($"Accent1 after reset: A={currentAccent1.A}, R={currentAccent1.R}, G={currentAccent1.G}, B={currentAccent1.B}");
        Console.WriteLine($"Default Accent1:      A={defaultAccent1.A}, R={defaultAccent1.R}, G={defaultAccent1.G}, B={defaultAccent1.B}");
        Console.WriteLine($"Reset successful: {currentAccent1.ToArgb() == defaultAccent1.ToArgb()}");

        // Save the workbook with the restored default theme
        workbook.Save("ResetToDefaultTheme.xlsx");
    }
}