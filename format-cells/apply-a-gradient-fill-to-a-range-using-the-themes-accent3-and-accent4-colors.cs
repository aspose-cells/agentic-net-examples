// Title: Apply a horizontal two‑color gradient using Accent3 & Accent4 theme colors in Aspose.Cells (.NET)
// Description: Creates a new workbook, extracts the Accent3 and Accent4 colors from the workbook theme, builds a Style with a horizontal two‑color gradient, applies it to the range B2:E6 via a StyleFlag, and saves the file as GradientThemeRange.xlsx using Aspose.Cells for C#.
// Keywords: Aspose.Cells gradient fill | theme colors Accent3 Accent4 | C# two‑color gradient | apply gradient to range Aspose.Cells | StyleFlag gradient .NET | horizontal gradient Aspose.Cells
// Common Searches: Aspose.Cells apply gradient fill to range | how to use workbook theme colors for gradient in C# | horizontal two‑color gradient Aspose.Cells example | retrieve Accent3 Accent4 colors Aspose.Cells | apply style with gradient to cell range .NET
// Developer Intent: Add a horizontal two‑color gradient to a specific cell range using the workbook’s Accent3 and Accent4 theme colors.
// Use Cases: Design a themed header row with a subtle horizontal gradient that follows the workbook’s color scheme. | Highlight a data block (e.g., B2:E6) with a gradient that automatically adapts when the workbook theme changes. | Generate reports where gradient fills stay consistent with corporate branding defined by theme accents.
// AI Prompts: Write C# code with Aspose.Cells to apply a vertical two‑color gradient using Accent1 and Accent2 theme colors. | Show how to change the gradient variant and direction for a range style in Aspose.Cells. | Provide an example of applying gradient fills to multiple non‑contiguous ranges using workbook theme colors in Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using AsposeRange = Aspose.Cells.Range;

// Creates a new workbook, extracts the Accent3 and Accent4 colors from the workbook theme, builds a Style with a horizontal two‑color gradient, applies it to the range B2:E6 via a StyleFlag, and saves the file as GradientThemeRange.xlsx using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Retrieve the theme colors Accent3 and Accent4 from the workbook's theme
            Color accent3 = workbook.GetThemeColor(ThemeColorType.Accent3);
            Color accent4 = workbook.GetThemeColor(ThemeColorType.Accent4);

            // Create a style and set a two‑color gradient using the theme colors
            Style gradientStyle = workbook.CreateStyle();
            gradientStyle.SetTwoColorGradient(
                accent3,                     // first theme color
                accent4,                     // second theme color
                GradientStyleType.Horizontal, // gradient direction
                1);                          // variant (1‑4)

            // Define the target range
            AsposeRange targetRange = worksheet.Cells.CreateRange("B2:E6");

            // Apply the gradient style to the range with a StyleFlag indicating which attributes to apply
            StyleFlag flag = new StyleFlag { All = true };
            targetRange.ApplyStyle(gradientStyle, flag);

            // Save the workbook
            workbook.Save("GradientThemeRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
