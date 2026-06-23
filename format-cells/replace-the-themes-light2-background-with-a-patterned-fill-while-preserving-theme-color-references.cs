using System.Drawing;
using Aspose.Cells;

class ReplaceLight2BackgroundWithPattern
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Create a style that uses a patterned fill
        Style style = workbook.CreateStyle();

        // Set the pattern type (e.g., diagonal stripe)
        style.Pattern = BackgroundType.DiagonalStripe;

        // Preserve the Light2 (Background2) theme reference for the background color
        style.BackgroundThemeColor = new ThemeColor(ThemeColorType.Background2, 0);

        // Optionally, use a theme accent for the foreground color of the pattern
        style.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);

        // Apply the style to a cell
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Patterned Light2");
        cell.SetStyle(style);

        // Save the workbook
        workbook.Save("PatternedLight2.xlsx", SaveFormat.Xlsx);
    }
}