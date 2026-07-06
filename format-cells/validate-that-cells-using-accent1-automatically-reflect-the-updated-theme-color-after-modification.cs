using System;
using System.Drawing;
using Aspose.Cells;

class ThemeColorUpdateDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Put a value into cell A1
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Accent1 Theme Test");

        // Create a style that uses Accent1 as the foreground theme color
        Style style = workbook.CreateStyle();
        style.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent1, 0); // No tint
        style.Pattern = BackgroundType.Solid;

        // Apply the style to the cell
        cell.SetStyle(style);

        // Display the original Accent1 theme color
        Color originalAccent1 = workbook.GetThemeColor(ThemeColorType.Accent1);
        Console.WriteLine("Original Accent1 theme color: " + originalAccent1);

        // Save the workbook before changing the theme (optional verification step)
        workbook.Save("BeforeThemeChange.xlsx");

        // Change the Accent1 theme color to a different color (e.g., Orange)
        workbook.SetThemeColor(ThemeColorType.Accent1, Color.Orange);

        // Retrieve and display the updated Accent1 theme color
        Color updatedAccent1 = workbook.GetThemeColor(ThemeColorType.Accent1);
        Console.WriteLine("Updated Accent1 theme color: " + updatedAccent1);

        // Verify that the cell still references Accent1 (its ThemeColor type remains Accent1)
        ThemeColor cellTheme = sheet.Cells["A1"].GetStyle().ForegroundThemeColor;
        Console.WriteLine("Cell foreground theme type: " + cellTheme.ColorType);
        Console.WriteLine("Cell foreground theme tint: " + cellTheme.Tint);

        // Save the workbook after the theme change
        workbook.Save("AfterThemeChange.xlsx");
    }
}