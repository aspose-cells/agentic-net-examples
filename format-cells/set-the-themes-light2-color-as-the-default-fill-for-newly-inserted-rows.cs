// Title: Set Light2 Theme Color as Default Row Fill When Inserting Rows – Aspose.Cells for .NET (C#)
// Description: Shows how to create a style that uses the workbook's Light2 (Background2) theme color with a solid fill, insert rows, and assign that style as the default row format so each newly added row automatically displays the Light2 background. The workbook is saved as RowsWithLight2ThemeFill.xlsx.
// Keywords: Aspose.Cells | C# | Light2 theme color | row default style | insert rows | theme background fill | Workbook theme | solid fill | Background2 | Style.SetStyle
// Common Searches: Aspose.Cells set row background to Light2 | C# apply workbook theme color to inserted rows | default row style Aspose.Cells | how to use theme colors in Aspose.Cells | insert rows with theme fill .NET
// Developer Intent: Apply the Light2 theme color as the default fill for rows inserted programmatically.
// Use Cases: Generate a report where every new data row automatically inherits the Light2 background for consistent visual styling. | Create a spreadsheet template that adds rows with a predefined theme‑based fill, eliminating manual formatting. | Build financial statements that dynamically insert rows while preserving the workbook’s Light2 theme color. | Design dashboards where inserted rows match the overall theme without extra code.
// AI Prompts: Modify the example to use a different theme color (e.g., Dark1) as the default fill for inserted rows. | Provide a version that applies the Light2 fill to a range of rows using a single style assignment instead of a loop. | Explain how to set the Light2 fill as the default style for all future rows added to the worksheet, including rows added after the initial insertion.

using System;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsThemeRowDemo
{
    // Shows how to create a style that uses the workbook's Light2 (Background2) theme color with a solid fill, insert rows, and assign that style as the default row format so each newly added row automatically displays the Light2 background. The workbook is saved as RowsWithLight2ThemeFill.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Prepare a style that uses the theme's Light2 (Background2) color
            // as the default fill for a row.
            // ------------------------------------------------------------
            Style light2Style = workbook.CreateStyle();

            // Use a solid fill pattern
            light2Style.Pattern = BackgroundType.Solid;

            // Set the background theme color to Light2 (Background2) with no tint
            // ThemeColorType.Background2 corresponds to the Light2 theme color.
            light2Style.BackgroundThemeColor = new ThemeColor(ThemeColorType.Background2, 0.0);

            // ------------------------------------------------------------
            // Insert a few rows and apply the prepared style as the default style
            // for each newly inserted row.
            // ------------------------------------------------------------
            // Insert three rows starting at row index 2 (third row, zero‑based)
            int insertIndex = 2;
            int rowsToInsert = 3;
            worksheet.Cells.InsertRows(insertIndex, rowsToInsert);

            // Apply the style to each inserted row
            for (int i = insertIndex; i < insertIndex + rowsToInsert; i++)
            {
                // Get the Row object
                Row row = worksheet.Cells.Rows[i];

                // Set the prepared style as the default style for this row
                row.SetStyle(light2Style);
            }

            // ------------------------------------------------------------
            // Optional: add some data to visualize the styled rows
            // ------------------------------------------------------------
            worksheet.Cells["A1"].PutValue("Header");
            worksheet.Cells["A3"].PutValue("Row with Light2 fill");
            worksheet.Cells["A4"].PutValue("Another Light2 row");
            worksheet.Cells["A5"].PutValue("Yet another Light2 row");

            // Save the workbook
            workbook.Save("RowsWithLight2ThemeFill.xlsx");
        }
    }
}
