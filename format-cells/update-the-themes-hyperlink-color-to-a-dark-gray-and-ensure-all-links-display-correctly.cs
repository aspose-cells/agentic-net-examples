using System;
using System.Drawing;
using Aspose.Cells;

class UpdateHyperlinkTheme
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Change the theme color for hyperlinks to a dark gray (RGB 64,64,64)
        workbook.SetThemeColor(ThemeColorType.Hyperlink, Color.FromArgb(64, 64, 64));

        // Add a sample hyperlink to demonstrate the new theme color
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Hyperlinks.Add("A1", 1, 1, "https://www.example.com");
        // Set display text for the hyperlink
        Hyperlink hyperlink = sheet.Hyperlinks[sheet.Hyperlinks.Count - 1];
        hyperlink.TextToDisplay = "Example Site";

        // Save the workbook with the updated hyperlink theme color
        workbook.Save("UpdatedHyperlinkTheme.xlsx");
    }
}