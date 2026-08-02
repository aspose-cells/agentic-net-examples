// Title: Apply a Horizontal Two‑Color Gradient with Theme Accent3 & Accent4 in Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, select a range (A1:C5), retrieve the workbook theme colors Accent3 and Accent4, build a style with a horizontal two‑color gradient, apply it to the range using a StyleFlag, and save the file as GradientFill_Accent3_Accent4.xlsx.
// Keywords: Aspose.Cells gradient fill | theme colors Accent3 Accent4 | horizontal two‑color gradient .NET | SetTwoColorGradient example | ApplyStyle range Aspose.Cells | Excel theme color fill C#
// Common Searches: Aspose.Cells apply gradient using theme colors | C# horizontal gradient Accent3 Accent4 | SetTwoColorGradient Aspose.Cells tutorial | How to use GetThemeColor for gradient fill | Apply style to cell range Aspose.Cells
// Developer Intent: Create and apply a horizontal two‑color gradient style to a specific cell range using the workbook’s Accent3 and Accent4 theme colors.
// Use Cases: Highlight header rows with a subtle themed gradient for visual separation. | Group sections in a financial report by shading backgrounds with a consistent brand palette. | Design reusable Excel templates where gradient fills automatically adapt to the workbook’s theme.
// AI Prompts: Generate code to change the gradient direction to vertical while still using Accent3 and Accent4. | Show how to apply a three‑color gradient using Accent2, Accent3, and Accent4 in Aspose.Cells. | Explain how to retrieve custom theme colors and use them for gradient fills in a .NET workbook.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, select a range (A1:C5), retrieve the workbook theme colors Accent3 and Accent4, build a style with a horizontal two‑color gradient, apply it to the range using a StyleFlag, and save the file as GradientFill_Accent3_Accent4.xlsx.
class GradientFillThemeExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range to which the gradient will be applied
            Aspose.Cells.Range range = sheet.Cells.CreateRange("A1:C5");

            // Retrieve the theme colors Accent3 and Accent4
            Color accent3 = workbook.GetThemeColor(ThemeColorType.Accent3);
            Color accent4 = workbook.GetThemeColor(ThemeColorType.Accent4);

            // Create a style and set its fill to a two‑color gradient
            Style style = workbook.CreateStyle();
            style.SetTwoColorGradient(accent3, accent4, GradientStyleType.Horizontal, 1);

            // Apply the style to the range (only cell shading)
            StyleFlag flag = new StyleFlag();
            flag.CellShading = true;
            range.ApplyStyle(style, flag);

            // Save the workbook
            workbook.Save("GradientFill_Accent3_Accent4.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
