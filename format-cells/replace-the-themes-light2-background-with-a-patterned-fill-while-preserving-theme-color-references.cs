// Title: Aspose.Cells .NET – Apply a Pattern Fill While Preserving the Light2 Theme Background
// Description: Shows how to create a workbook, define a diagonal‑stripe pattern style, set a red foreground, keep the Light2 theme background using BackgroundThemeColor, apply the style to cell B2, and save the file as XLSX.
// Keywords: Aspose.Cells | .NET | pattern fill | BackgroundThemeColor | Light2 theme | ThemeColorType.Background2 | diagonal stripe | cell style | workbook theme | preserve theme colors
// Common Searches: Aspose.Cells pattern fill with Light2 theme background | set BackgroundThemeColor for patterned style .NET | how to keep theme colors when using pattern fill Aspose.Cells | apply diagonal stripe pattern to a cell in C# | preserve workbook theme while styling cells
// Developer Intent: Add a patterned fill to a cell while retaining the Light2 theme background reference.
// Use Cases: Design reports that combine visual patterns with theme‑based colors for consistent branding. | Create reusable templates where pattern fills respect the workbook’s Light2 theme across multiple sheets. | Enhance accessibility by using theme colors for background while highlighting cells with patterned overlays.
// AI Prompts: Generate C# code that applies a diagonal‑stripe pattern to a cell and uses BackgroundThemeColor to keep the Light2 theme background in Aspose.Cells. | Explain how to combine ForegroundColor and BackgroundThemeColor to create a theme‑aware patterned fill in a .NET workbook. | Show an example of applying the same Light2‑based pattern style to a range of cells using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsThemePatternDemo
{
    // Shows how to create a workbook, define a diagonal‑stripe pattern style, set a red foreground, keep the Light2 theme background using BackgroundThemeColor, apply the style to cell B2, and save the file as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Define a style that uses a pattern fill
            Style patternStyle = workbook.CreateStyle();

            // Set the pattern type (e.g., diagonal stripe)
            patternStyle.Pattern = BackgroundType.DiagonalStripe;

            // Set the foreground color of the pattern (regular Color)
            patternStyle.ForegroundColor = Color.Red;

            // Preserve the theme reference for the background by using BackgroundThemeColor
            // ThemeColorType.Background2 corresponds to the Light2 background color of the theme
            patternStyle.BackgroundThemeColor = new ThemeColor(ThemeColorType.Background2, 0);

            // Apply the style to a cell
            Cell targetCell = sheet.Cells["B2"];
            targetCell.PutValue("Pattern with Light2 background");
            targetCell.SetStyle(patternStyle);

            // Save the workbook (lifecycle: save)
            workbook.Save("ThemePatternFillDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
