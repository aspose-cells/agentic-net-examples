using System;
using System.Drawing;
using Aspose.Cells;

class UpdateHyperlinkThemeColor
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Define a new shade of blue for the Hyperlink theme color
        Color newBlue = Color.FromArgb(0, 112, 192); // custom blue

        // Update the Hyperlink theme color
        workbook.SetThemeColor(ThemeColorType.Hyperlink, newBlue);

        // Demonstrate the theme color by applying it to a cell's font
        Worksheet sheet = workbook.Worksheets[0];
        Cell cell = sheet.Cells["A1"];
        cell.PutValue("Hyperlink Example");
        Style style = workbook.CreateStyle();
        style.Font.ThemeColor = new ThemeColor(ThemeColorType.Hyperlink, 0);
        cell.SetStyle(style);

        // Save the workbook
        workbook.Save("HyperlinkThemeUpdated.xlsx");
    }
}