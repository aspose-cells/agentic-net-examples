using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemeColorValidationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Apply Accent1 theme color to cell A1 (no tint)
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Accent1 Theme Test");

            Style style = workbook.CreateStyle();
            style.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0);
            style.Pattern = BackgroundType.Solid;
            cell.SetStyle(style);

            // Save the initial workbook (optional, demonstrates lifecycle rule: save)
            workbook.Save("InitialThemeDemo.xlsx");

            // Change the Accent1 theme color to a new concrete color (e.g., Red)
            workbook.SetThemeColor(ThemeColorType.Accent1, Color.Red);

            // Retrieve the updated theme color from the workbook
            Color updatedThemeColor = workbook.GetThemeColor(ThemeColorType.Accent1);
            Console.WriteLine($"Updated Theme Color for Accent1: {updatedThemeColor}");

            // Verify that the cell still references Accent1 (theme type unchanged)
            Style refreshedStyle = cell.GetStyle();
            ThemeColor cellTheme = refreshedStyle.ForegroundThemeColor;
            Console.WriteLine($"Cell A1 Theme Type: {cellTheme.ColorType}, Tint: {cellTheme.Tint}");

            // Save the workbook after theme modification (lifecycle rule: save)
            workbook.Save("UpdatedThemeDemo.xlsx");

            // Indicate completion
            Console.WriteLine("Validation completed. Check the generated Excel files.");
        }
    }
}