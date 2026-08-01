// Title: C# – Lighten Theme Background2 (Light2) and Keep Dark Text Using Aspose.Cells
// Description: Demonstrates how to set the Light2 (Background2) theme color to a lighter shade, apply a 50 % tint for a solid fill, and use a dark Text1 font so the text remains readable. The example creates a workbook, styles cell A1, and saves the file as UpdatedThemeLight2.xlsx.
// Keywords: Aspose.Cells C# theme color | Background2 Light2 tint | set theme background color .NET | dark font on light background Excel | Workbook.SetThemeColor example | theme color styling Aspose.Cells
// Common Searches: how to change Light2 theme color in Aspose.Cells | apply tint to Background2 theme color C# | make dark text readable on light theme background Excel | Aspose.Cells set theme background2 to light gray | C# example for theme color tint in Excel workbook
// Developer Intent: Adjust the Light2 (Background2) theme color to a lighter shade while preserving dark text contrast.
// Use Cases: Create a report workbook with a light gray Background2 applied to selected cells for a clean visual style. | Apply a 50 % lighten tint to the Background2 theme color on specific cells while keeping the font color dark for optimal contrast. | Generate Excel files where the Light2 theme background is brightened to improve on‑screen and printed readability of dark text.
// AI Prompts: Generate C# code that changes the Light2 (Background2) theme color to a custom RGB value and applies it to a range of cells with Aspose.Cells. | Provide an Aspose.Cells snippet that calculates an optimal tint for Background2 based on existing theme colors to ensure sufficient contrast with dark Text1 font.

using System;
using System.Drawing;
using Aspose.Cells;

namespace ThemeUpdateExample
{
    // Demonstrates how to set the Light2 (Background2) theme color to a lighter shade, apply a 50 % tint for a solid fill, and use a dark Text1 font so the text remains readable. The example creates a workbook, styles cell A1, and saves the file as UpdatedThemeLight2.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the base theme color for Background2 (Light2) to a light gray
            workbook.SetThemeColor(ThemeColorType.Background2, Color.LightGray);

            // Prepare a cell to demonstrate the updated background and dark text
            Cell cell = worksheet.Cells["A1"];
            cell.PutValue("Dark text on lighter background");

            // Create a style that uses the Background2 theme color with a positive tint (lighten)
            Style style = workbook.CreateStyle();
            style.Pattern = BackgroundType.Solid;
            // Apply a 50% lighten tint to the background theme color
            style.BackgroundThemeColor = new ThemeColor(ThemeColorType.Background2, 0.5);
            // Ensure the font uses a dark theme color (Text1) without tint
            style.Font.ThemeColor = new ThemeColor(ThemeColorType.Text1, 0.0);

            // Apply the style to the cell
            cell.SetStyle(style);

            // Save the workbook
            workbook.Save("UpdatedThemeLight2.xlsx");
        }
    }
}
