// Title: Apply Accent1 Theme Color as Cell Background in Aspere.Cells (C#)
// Description: Creates a new workbook, defines a solid‑fill style whose background uses the Accent1 theme color, builds a range covering A1:D10 on the first worksheet, applies only the shading style to that range, and saves the file as an .xlsx document.
// Keywords: Aspose.Cells C# theme background | Accent1 cell fill | set theme color for range | solid fill style Aspose.Cells | Excel theme color programmatically
// Common Searches: Aspose.Cells apply theme color to cell background | C# set Accent1 as fill color in Excel | how to use ThemeColorType.Accent1 in Aspose.Cells | apply solid background to a range with Aspose.Cells .NET
// Developer Intent: Color the background of a specific range on the first worksheet using the workbook’s Accent1 theme color.
// Use Cases: Highlight header rows with the primary theme shade for consistent branding. | Mark important data sections in generated reports using the Accent1 background. | Create reusable Excel templates where key areas automatically adopt the Accent1 fill.
// AI Prompts: Show how to use Accent2 as a solid background for a range with Aspose.Cells. | Provide C# code that adds a 20 % tint to the Accent1 background for cells A1:D10. | Explain how to define a reusable Accent1 background style and apply it across multiple worksheets.

using System;
using Aspose.Cells;

namespace AsposeCellsThemeDemo
{
    // Creates a new workbook, defines a solid‑fill style whose background uses the Accent1 theme color, builds a range covering A1:D10 on the first worksheet, applies only the shading style to that range, and saves the file as an .xlsx document.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Create a style that uses the Accent1 theme color as background
                Style accentStyle = workbook.CreateStyle();
                accentStyle.Pattern = BackgroundType.Solid; // solid fill
                accentStyle.BackgroundThemeColor = new ThemeColor(ThemeColorType.Accent1, 0.0); // Accent1, no tint

                // Define the range (A1:D10)
                CellArea area = new CellArea
                {
                    StartRow = 0,
                    StartColumn = 0,
                    EndRow = 9,
                    EndColumn = 3
                };

                // Calculate rows and columns for the range
                int totalRows = area.EndRow - area.StartRow + 1;
                int totalColumns = area.EndColumn - area.StartColumn + 1;

                // Create the Aspose.Cells range (avoid conflict with System.Range)
                Aspose.Cells.Range range = cells.CreateRange(area.StartRow, area.StartColumn, totalRows, totalColumns);

                // Apply only the background style
                StyleFlag flag = new StyleFlag { CellShading = true };
                range.ApplyStyle(accentStyle, flag);

                // Save the workbook
                workbook.Save("Accent1BackgroundDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
