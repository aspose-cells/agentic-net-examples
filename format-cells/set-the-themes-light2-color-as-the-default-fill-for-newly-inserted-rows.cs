// Title: Aspose.Cells for .NET – Apply Light2 Theme Color as Default Fill for Inserted Rows
// Description: Demonstrates how to create a style that uses the workbook's Light2 (Background2) theme color, insert rows at a specific index, and assign the style to each new Row so that any later cells inherit a solid Light2 background. The workbook is then saved as an .xlsx file.
// Keywords: Aspose.Cells | C# | .NET | theme color | Light2 | Background2 | row style | insert rows | default fill | solid background | Style.SetStyle | Row.SetStyle
// Common Searches: set theme Light2 fill for new rows Aspose.Cells | default row background using workbook theme .NET | apply theme color to inserted rows C# | how to use Background2 theme color in Aspose.Cells | row style inheritance after InsertRows
// Developer Intent: Apply the workbook's Light2 theme color as the default background for rows that are programmatically inserted.
// Use Cases: Create a template where any added rows automatically match the secondary theme color, ensuring visual consistency. | Generate dynamic reports that insert data rows while preserving the workbook’s theme‑based styling without setting each cell individually. | Build financial models where new rows inherit the Background2 fill, simplifying maintenance of a uniform look across changing data sections.
// AI Prompts: Show how to switch the style to use the Dark1 theme color for newly inserted rows. | Provide code to apply the same Light2 theme style to an entire column after rows are added. | Explain how to keep the row style when copying rows to another worksheet with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeRowDemo
{
    // Demonstrates how to create a style that uses the workbook's Light2 (Background2) theme color, insert rows at a specific index, and assign the style to each new Row so that any later cells inherit a solid Light2 background. The workbook is then saved as an .xlsx file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Define a style that uses the theme's Light2 color (Background2)
            // ------------------------------------------------------------
            Style themeStyle = workbook.CreateStyle();
            // Use the theme color Background2 (Light2) with no tint
            themeStyle.BackgroundThemeColor = new ThemeColor(ThemeColorType.Background2, 0);
            // Set a solid fill pattern so the background color is applied
            themeStyle.Pattern = BackgroundType.Solid;

            // ------------------------------------------------------------
            // Insert new rows (for example, insert 3 rows at index 2)
            // ------------------------------------------------------------
            int insertIndex = 2;   // zero‑based row index where rows will be inserted
            int rowCount = 3;      // number of rows to insert
            worksheet.Cells.InsertRows(insertIndex, rowCount);

            // ------------------------------------------------------------
            // Apply the theme style as the default style for each newly inserted row
            // ------------------------------------------------------------
            for (int i = insertIndex; i < insertIndex + rowCount; i++)
            {
                // Get the Row object
                Row row = worksheet.Cells.Rows[i];
                // Set the style; this becomes the default style for cells added later in this row
                row.SetStyle(themeStyle);
            }

            // ------------------------------------------------------------
            // (Optional) Add some data to the new rows to demonstrate the style
            // ------------------------------------------------------------
            for (int i = insertIndex; i < insertIndex + rowCount; i++)
            {
                worksheet.Cells[i, 0].PutValue($"Row {i + 1}");
                worksheet.Cells[i, 1].PutValue(i * 10);
            }

            // Save the workbook
            workbook.Save("ThemeLight2RowStyle.xlsx");
        }
    }
}
