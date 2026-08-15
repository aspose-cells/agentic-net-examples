// Title: Aspose.Cells .NET: Create an Excel Theme‑Color Preview Grid
// Description: C# example that builds a new Workbook, lists every ThemeColorType in a 4‑column grid, applies the same theme color to the font and a light‑tinted fill, and saves the sheet as ThemeColorPreview.xlsx.
// Keywords: Aspose.Cells | .NET | C# | ThemeColorType | Excel theme colors | preview grid | visual palette | sample code | cell styling | foreground background theme
// Common Searches: Aspose.Cells theme color preview example | how to display all ThemeColorType values in Excel | C# code to create a theme‑color palette with Aspose.Cells | visualize Excel theme colors programmatically | generate theme color grid using Aspose.Cells .NET
// Developer Intent: Produce an Excel worksheet that visually showcases each ThemeColorType with matching font and background colors.
// Use Cases: Quick reference for designers to see how theme colors appear as text and fill. | Test workbook for verifying theme‑color rendering across Office versions. | Printable color palette for style guides or documentation.
// AI Prompts: Write C# code with Aspose.Cells that creates a 3‑column theme‑color grid and exports it to PDF. | Add conditional formatting to the preview grid so Accent colors receive a thick border. | Modify the sample to include custom tints for each theme color and generate a CSV summary.

using System;
using Aspose.Cells;

namespace AsposeCellsThemeColorPreview
{
    // C# example that builds a new Workbook, lists every ThemeColorType in a 4‑column grid, applies the same theme color to the font and a light‑tinted fill, and saves the sheet as ThemeColorPreview.xlsx.
    public class ThemeColorGridDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Define all ThemeColorType values to preview
                ThemeColorType[] themeTypes = new ThemeColorType[]
                {
                    ThemeColorType.Background1,
                    ThemeColorType.Text1,
                    ThemeColorType.Background2,
                    ThemeColorType.Text2,
                    ThemeColorType.Accent1,
                    ThemeColorType.Accent2,
                    ThemeColorType.Accent3,
                    ThemeColorType.Accent4,
                    ThemeColorType.Accent5,
                    ThemeColorType.Accent6,
                    ThemeColorType.Hyperlink,
                    ThemeColorType.FollowedHyperlink,
                    ThemeColorType.StyleColor
                };

                // Layout settings: 4 columns per row
                int columns = 4;
                int startRow = 0;
                int startColumn = 0;

                for (int i = 0; i < themeTypes.Length; i++)
                {
                    int row = startRow + i / columns;
                    int col = startColumn + i % columns;

                    // Put the name of the theme color in the cell
                    Cell cell = cells[row, col];
                    cell.PutValue(themeTypes[i].ToString());

                    // Create a style for the cell
                    Style style = workbook.CreateStyle();

                    // Use the theme color as the foreground (font) color
                    style.Font.ThemeColor = new ThemeColor(themeTypes[i], 0.0);
                    style.Font.Size = 12;
                    style.Font.IsBold = true;

                    // Set a background theme color with a light tint for visibility
                    style.ForegroundThemeColor = new ThemeColor(themeTypes[i], 0.5);
                    style.Pattern = BackgroundType.Solid;

                    // Apply the style to the cell
                    cell.SetStyle(style);
                }

                // Save the workbook
                workbook.Save("ThemeColorPreview.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ThemeColorGridDemo.Run();
        }
    }
}
