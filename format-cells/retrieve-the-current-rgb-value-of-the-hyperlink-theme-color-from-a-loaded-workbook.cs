using System;
using System.Drawing;
using Aspose.Cells;

class RetrieveHyperlinkThemeColor
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Retrieve the theme color for Hyperlink
        Color hyperlinkColor = workbook.GetThemeColor(ThemeColorType.Hyperlink);

        // Display the RGB components of the Hyperlink theme color
        Console.WriteLine($"Hyperlink Theme Color - R:{hyperlinkColor.R}, G:{hyperlinkColor.G}, B:{hyperlinkColor.B}");

        // Save the workbook (no modifications made, just to follow lifecycle rules)
        workbook.Save("output.xlsx");
    }
}