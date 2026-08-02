// Title: C# Sample: Generate an Excel ThemeColorType Preview Grid with Aspose.Cells
// Description: A C# example that creates an Excel workbook, lists each ThemeColorType enum entry, and shows a solid‑filled cell with the matching theme color. The sheet includes a header, auto‑fits columns, and is saved as ThemeColorsPreview.xlsx.
// Keywords: Aspose.Cells | ThemeColorType | theme color preview | C# Excel sample | .NET Aspose.Cells tutorial | generate Excel palette | solid fill style | auto fit columns | workbook creation | visual theme colors
// Common Searches: Aspose.Cells C# list all ThemeColorType colors | how to preview Excel theme colors programmatically | sample code for theme color palette in .NET | create a theme color reference sheet with Aspose.Cells | display ThemeColorType enum values in Excel
// Developer Intent: Produce an Excel worksheet that enumerates every ThemeColorType and displays a colored cell preview for each entry.
// Use Cases: Design teams can reference a complete set of built‑in theme colors. | Documentation writers can embed a visual palette of Excel theme colors. | Automated tests can verify correct mapping of theme colors to cells.
// AI Prompts: Write C# code using Aspose.Cells that adds a row for each ThemeColorType, fills a cell with the corresponding theme color, and saves the file. | Modify the example to show three tint variations (e.g., -0.5, 0, 0.5) for each theme color in separate columns. | Create a method that returns a dictionary mapping ThemeColorType names to their RGB values based on the generated workbook.

using System;
using Aspose.Cells;
using System.Drawing;

namespace AsposeCellsThemeColorPreview
{
    // A C# example that creates an Excel workbook, lists each ThemeColorType enum entry, and shows a solid‑filled cell with the matching theme color. The sheet includes a header, auto‑fits columns, and is saved as ThemeColorsPreview.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Header row
            cells["A1"].PutValue("Theme Color Type");
            cells["B1"].PutValue("Preview");

            // Loop through all ThemeColorType enum values
            int row = 1; // start from second row (index 1)
            foreach (ThemeColorType type in Enum.GetValues(typeof(ThemeColorType)))
            {
                // Write the enum name
                cells[row, 0].PutValue(type.ToString());

                // Create a style that uses the theme color as foreground (solid fill)
                Style style = workbook.CreateStyle();
                style.ForegroundThemeColor = new ThemeColor(type, 0.0); // no tint
                style.Pattern = BackgroundType.Solid;

                // Apply the style to the preview cell
                cells[row, 1].PutValue(" "); // placeholder text
                cells[row, 1].SetStyle(style);

                row++;
            }

            // Adjust column widths for better visibility
            worksheet.AutoFitColumn(0);
            worksheet.AutoFitColumn(1);

            // Save the workbook
            workbook.Save("ThemeColorsPreview.xlsx");
        }
    }
}
