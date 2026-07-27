// Title: Aspose.Cells for .NET – Replace Light2 Theme Background with a Pattern Fill (C#)
// Description: Demonstrates how to change the Light2 (Background2) theme color of a cell to a patterned fill such as DiagonalStripe while keeping theme references. The example creates a workbook, adds text to A1, sets BackgroundThemeColor and ForegroundThemeColor, applies the style, and saves the file as ThemePatternFill.xlsx.
// Keywords: Aspose.Cells | C# | .NET | Excel pattern fill | Light2 theme background | BackgroundThemeColor | ForegroundThemeColor | BackgroundType DiagonalStripe | theme color fill | replace solid background with pattern | Excel styling example
// Common Searches: Aspose.Cells replace Light2 background with pattern | C# set patterned fill using theme colors in Excel | BackgroundThemeColor DiagonalStripe Aspose.Cells | How to use theme colors for pattern fill in .NET workbook | Excel pattern fill preserving theme colors Aspose
// Developer Intent: Apply a patterned fill to a cell while preserving the workbook’s Light2 (Background2) theme color using Aspose.Cells for .NET.
// Use Cases: Design reports where theme colors are maintained but cells need visual distinction via patterns. | Create Excel templates that replace solid theme backgrounds with diagonal stripe or other patterns. | Generate export files that stay consistent with corporate Light2 theme while adding patterned styling.
// AI Prompts: Show C# code to replace the Light2 theme background with a diagonal stripe pattern using Aspose.Cells. | How can I apply a pattern fill while keeping theme color references in an Aspose.Cells workbook? | Provide an Aspose.Cells example that sets BackgroundThemeColor and ForegroundThemeColor for a patterned cell style.

using System;
using System.Drawing;
using Aspose.Cells;

// Demonstrates how to change the Light2 (Background2) theme color of a cell to a patterned fill such as DiagonalStripe while keeping theme references. The example creates a workbook, adds text to A1, sets BackgroundThemeColor and ForegroundThemeColor, applies the style, and saves the file as ThemePatternFill.xlsx.
class ReplaceThemeBackground
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Put a sample value in a cell that will use the themed background
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue("Themed Pattern Fill");

        // Create a new style
        Style style = workbook.CreateStyle();

        // Set a pattern type (e.g., DiagonalStripe) instead of a solid background
        style.Pattern = BackgroundType.DiagonalStripe;

        // Preserve the original theme background color (Light2 corresponds to Background2)
        style.BackgroundThemeColor = new ThemeColor(ThemeColorType.Background2, 0);

        // Optionally set a foreground theme color for the pattern (e.g., Accent1)
        style.ForegroundThemeColor = new ThemeColor(ThemeColorType.Accent1, 0);

        // Apply the style to the cell
        cell.SetStyle(style);

        // Save the workbook
        workbook.Save("ThemePatternFill.xlsx", SaveFormat.Xlsx);
    }
}
