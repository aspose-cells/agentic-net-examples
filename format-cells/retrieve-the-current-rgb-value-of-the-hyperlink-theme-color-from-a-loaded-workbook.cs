using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsHyperlinkThemeColorDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Retrieve the theme color for Hyperlink
            Color hyperlinkColor = workbook.GetThemeColor(ThemeColorType.Hyperlink);

            // Output the RGB components of the retrieved color
            Console.WriteLine($"Hyperlink Theme Color - A:{hyperlinkColor.A}, R:{hyperlinkColor.R}, G:{hyperlinkColor.G}, B:{hyperlinkColor.B}");
        }
    }
}