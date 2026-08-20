// Title: Apply Accent1 Theme Color as Cell Background in the First Worksheet (C# Aspose.Cells)
// Description: Creates a new workbook, accesses the first worksheet, defines a solid‑fill style using the Accent1 theme color, applies it to cells A1:D10, and saves the file as Accent1Background.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | theme color background | Accent1 | solid fill style | set cell background | ThemeColor | Excel styling | Workbook | Worksheet
// Common Searches: Aspose.Cells set Accent1 background | C# apply theme color to cell range Aspose.Cells | How to use ThemeColor in Aspose.Cells .NET | Set solid fill background with Accent1 in Excel using Aspose | Apply theme color to multiple cells Aspose.Cells
// Developer Intent: Apply the Accent1 theme color as a solid‑fill background to a range of cells in the first worksheet.
// Use Cases: Highlight header rows (e.g., A1:D1) with an Accent1 background to make titles stand out. | Create a themed data table (A1:D10) where all cells share the same Accent1 fill for visual consistency. | Build a reusable workbook template that automatically styles title cells with the Accent1 background.
// AI Prompts: Generate C# code that applies the Accent2 theme color as a gradient fill to cells B2:E5 using Aspose.Cells. | Show how to adjust the tint of an Accent1 background style to a lighter shade in Aspose.Cells. | Provide an example of applying a custom theme color to the entire worksheet background with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsThemeDemo
{
    // Creates a new workbook, accesses the first worksheet, defines a solid‑fill style using the Accent1 theme color, applies it to cells A1:D10, and saves the file as Accent1Background.xlsx with Aspose.Cells for .NET.
    public class ApplyAccent1Background
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a style with solid fill and set its background to the Accent1 theme color
            Style accentStyle = workbook.CreateStyle();
            accentStyle.Pattern = BackgroundType.Solid;
            // ThemeColor constructor: (ThemeColorType, tint). Tint 0 means original color.
            accentStyle.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);

            // Define the range of cells to which the style will be applied (A1:D10)
            int startRow = 0;
            int endRow = 9;
            int startColumn = 0;
            int endColumn = 3;

            // Apply the style to each cell in the defined range
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startColumn; col <= endColumn; col++)
                {
                    Cell cell = worksheet.Cells[row, col];
                    cell.SetStyle(accentStyle);
                }
            }

            // Save the workbook
            string outputPath = "Accent1Background.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
