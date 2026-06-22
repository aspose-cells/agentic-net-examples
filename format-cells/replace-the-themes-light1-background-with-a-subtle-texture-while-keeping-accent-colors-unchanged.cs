using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeTextureDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Retrieve the current Background1 (Light1) theme color
            Color background1Color = workbook.GetThemeColor(ThemeColorType.Background1);

            // Get the default style of the workbook
            Style defaultStyle = workbook.DefaultStyle;

            // Apply a subtle texture (Gray25 pattern) using the theme's Background1 color as the foreground
            // and a light color (e.g., White) as the background of the pattern.
            defaultStyle.SetPatternColor(BackgroundType.Gray25, background1Color, Color.White);

            // Assign the modified style back as the workbook's default style
            workbook.DefaultStyle = defaultStyle;

            // (Optional) Demonstrate the effect on a sample cell
            Worksheet sheet = workbook.Worksheets[0];
            Cell sampleCell = sheet.Cells["A1"];
            sampleCell.PutValue("Cell with Light1 texture");
            sampleCell.SetStyle(defaultStyle);

            // Save the workbook
            workbook.Save("ThemeWithLight1Texture.xlsx");
        }
    }
}