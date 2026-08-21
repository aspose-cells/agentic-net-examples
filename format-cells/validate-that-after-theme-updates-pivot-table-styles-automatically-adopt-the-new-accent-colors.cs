// Title: Verify Pivot Table Style Updates with New Theme Accent Colors using Aspose.Cells for .NET
// Description: Shows how to create a workbook, add a pivot table that uses a built‑in style tied to theme accents, change Accent1 and Accent2 with SetThemeColor, refresh the pivot, and programmatically confirm the cell’s ForegroundThemeColor now reflects the updated accent.
// Keywords: Aspose.Cells | C# | pivot table | theme colors | SetThemeColor | RefreshPivotTables | theme accent validation | PivotTableStyleMedium10 | Excel theme automation | cell style theme color
// Common Searches: how to update pivot table theme colors in Aspose.Cells | Aspose.Cells verify pivot table uses new accent color | refresh pivot tables after changing workbook theme | C# check ForegroundThemeColor of pivot table cell | SetThemeColor effect on existing pivot tables
// Developer Intent: Confirm that a pivot table’s formatting automatically follows changes to workbook theme accent colors after invoking SetThemeColor and RefreshPivotTables.
// Use Cases: Automated testing of theme propagation to existing pivot tables in CI pipelines. | Generating reports where custom theme colors must be applied without rebuilding pivot tables. | Validating that built‑in pivot styles correctly reference updated Accent1/Accent2 values.
// AI Prompts: Write C# code with Aspose.Cells that changes Accent1 to red, refreshes a pivot table, and checks that a data cell now uses the red accent. | Explain step‑by‑step how RefreshPivotTables re‑evaluates theme colors for pivot tables in Aspose.Cells. | Create a unit test that asserts a pivot table cell’s ForegroundThemeColor equals ThemeColorType.Accent1 after a theme color change.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsThemePivotValidation
{
    // Shows how to create a workbook, add a pivot table that uses a built‑in style tied to theme accents, change Accent1 and Accent2 with SetThemeColor, refresh the pivot, and programmatically confirm the cell’s ForegroundThemeColor now reflects the updated accent.
    class Program
    {
        static void Main()
        {
            // -------------------- Create a new workbook --------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Region");
            sheet.Cells["C1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("Laptop");
            sheet.Cells["B2"].PutValue("North");
            sheet.Cells["C2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("Desktop");
            sheet.Cells["B3"].PutValue("South");
            sheet.Cells["C3"].PutValue(800);

            sheet.Cells["A4"].PutValue("Tablet");
            sheet.Cells["B4"].PutValue("East");
            sheet.Cells["C4"].PutValue(500);

            // -------------------- Add a pivot table --------------------
            int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivot.AddFieldToArea(PivotFieldType.Row, 0);      // Product
            pivot.AddFieldToArea(PivotFieldType.Column, 1);   // Region
            pivot.AddFieldToArea(PivotFieldType.Data, 2);     // Sales

            // Apply a built‑in pivot table style that uses theme accent colors
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium10;

            // -------------------- Save the workbook before theme change (optional) --------------------
            workbook.Save("BeforeThemeChange.xlsx");

            // -------------------- Update theme accent colors --------------------
            // Change Accent1 to Red and Accent2 to Green
            workbook.SetThemeColor(ThemeColorType.Accent1, Color.Red);
            workbook.SetThemeColor(ThemeColorType.Accent2, Color.Green);

            // Refresh pivot tables so they re‑evaluate the theme
            sheet.RefreshPivotTables();

            // -------------------- Validate that the pivot table adopted the new theme --------------------
            // Get the first data cell of the pivot table
            CellArea dataArea = pivot.DataBodyRange;
            int firstDataRow = dataArea.StartRow;
            int firstDataColumn = dataArea.StartColumn;
            Style dataCellStyle = sheet.Cells[firstDataRow, firstDataColumn].GetStyle();

            // The style should reference a theme color (e.g., Accent1). Verify the ThemeColorType.
            ThemeColor themeColor = dataCellStyle.ForegroundThemeColor;
            if (themeColor != null && themeColor.ColorType == ThemeColorType.Accent1)
            {
                Console.WriteLine("Pivot table cell uses Theme Accent1 as expected.");
                // Retrieve the actual color applied after the theme change
                Color actualColor = dataCellStyle.ForegroundColor;
                Console.WriteLine($"Actual foreground color after theme update: {actualColor}");
            }
            else
            {
                Console.WriteLine("Pivot table cell does NOT use the expected theme accent color.");
            }

            // -------------------- Save the final workbook --------------------
            workbook.Save("AfterThemeChange.xlsx");
        }
    }
}
