using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemeRetentionValidation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define a theme color (Accent1) as Red
            workbook.SetThemeColor(ThemeColorType.Accent1, Color.Red);

            // Apply the theme color to a cell's font
            Cell themedCell = sheet.Cells["A1"];
            themedCell.PutValue("Themed Text");
            Style themedStyle = workbook.CreateStyle();
            themedStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0); // No tint
            themedCell.SetStyle(themedStyle);

            // Capture the actual RGB color that the theme resolves to
            Color resolvedColor = workbook.GetThemeColor(ThemeColorType.Accent1);
            Console.WriteLine($"Resolved theme color (Accent1) before change: {resolvedColor}");

            // Freeze the visual appearance by converting the theme color to a solid RGB color
            Style frozenStyle = workbook.CreateStyle();
            frozenStyle.Font.Color = resolvedColor; // Direct RGB color, no theme reference
            frozenStyle.Font.Size = 12;
            themedCell.SetStyle(frozenStyle);

            // Change the theme color to Blue (simulating theme removal/modification)
            workbook.SetThemeColor(ThemeColorType.Accent1, Color.Blue);
            Color newResolved = workbook.GetThemeColor(ThemeColorType.Accent1);
            Console.WriteLine($"Resolved theme color (Accent1) after change: {newResolved}");

            // Verify that the cell's font color remains the original Red
            Color cellFontColor = themedCell.GetStyle().Font.Color;
            Console.WriteLine($"Cell font color after theme change: {cellFontColor}");

            // Save the workbook to inspect the result manually if needed
            workbook.Save("ThemeRetentionValidation.xlsx");
        }
    }
}