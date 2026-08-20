// Title: Aspose.Cells .NET: Light2 Theme Background – Lighter Tint with Dark Text
// Description: Creates a new workbook, defines a style that applies a 50 % lighter Background2 (Light2) theme color, sets the font to the dark Text1 theme color with bold formatting, applies the style to a range of cells, and saves the file as UpdatedLight2Theme.xlsx.
// Keywords: Aspose.Cells | .NET | C# | Light2 theme | Background2 tint | theme color tint | dark text readability | Excel style programmatically | theme color adjustment | Workbook styling
// Common Searches: Aspose.Cells Light2 background lighter | set Background2 tint in .NET | dark font on light theme Aspose.Cells | adjust theme colors programmatically Excel | C# change workbook theme background
// Developer Intent: Programmatically lighten the Light2 (Background2) theme color of a workbook and keep the text dark for optimal contrast using Aspose.Cells for .NET.
// Use Cases: Design report headers with a subtle Light2 background while preserving bold, dark text for clarity. | Prepare printable spreadsheets where a lighter theme background reduces ink usage without sacrificing legibility. | Apply a custom style to specific rows or columns that requires a lighter background shade and high‑contrast dark font.
// AI Prompts: Show how to apply a 50 % tint to the Background2 theme color in Aspose.Cells C#. | Provide code that sets a lighter Light2 background and uses Text1 dark font with bold styling. | Explain the steps to modify theme colors in an Aspose.Cells workbook while ensuring text readability.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeUpdate
{
    // Creates a new workbook, defines a style that applies a 50 % lighter Background2 (Light2) theme color, sets the font to the dark Text1 theme color with bold formatting, applies the style to a range of cells, and saves the file as UpdatedLight2Theme.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Define a style that uses the Light2 (Background2) theme color with a lighter tint
            Style lightStyle = workbook.CreateStyle();
            // Set solid fill pattern
            lightStyle.Pattern = BackgroundType.Solid;
            // Apply Background2 theme color with a positive tint (e.g., 0.5 makes it 50% lighter)
            lightStyle.BackgroundThemeColor = new ThemeColor(ThemeColorType.Background2, 0.5);

            // Ensure dark text is readable on the lighter background
            // Use Text1 theme color (typically dark) with no tint adjustment
            lightStyle.Font.ThemeColor = new ThemeColor(ThemeColorType.Text1, 0.0);
            lightStyle.Font.IsBold = true;

            // Apply the style to a range of cells for demonstration
            for (int row = 0; row < 5; row++)
            {
                cells[row, 0].PutValue($"Row {row + 1}");
                cells[row, 0].SetStyle(lightStyle);
            }

            // Save the workbook
            workbook.Save("UpdatedLight2Theme.xlsx");
        }
    }
}
