// Title: Validate PivotTable Style Adopts New Workbook Theme Accent Using Aspose.Cells for .NET
// Description: This C# example creates a workbook, builds a pivot table with a built‑in style that uses theme accent colors, records the foreground ARGB of a data cell, changes the Accent1 theme color to red via Workbook.SetThemeColor, refreshes the pivot tables, re‑reads the cell color, and confirms the style automatically reflects the updated theme.
// Keywords: Aspose.Cells | PivotTable | SetThemeColor | RefreshPivotTables | .NET | C# | Excel theme accent | style validation | automated testing | CI pipeline
// Common Searches: Aspose.Cells change pivot table theme color | RefreshPivotTables after SetThemeColor | Validate Excel theme accent update in .NET | PivotTable style follows workbook theme change | How to test theme‑driven styling with Aspose.Cells
// Developer Intent: Verify that a pivot table’s style automatically updates when the workbook’s theme accent color is changed.
// Use Cases: Automated verification of theme‑dependent pivot table styling in reporting solutions. | CI/CD checks to ensure Excel files generated with Aspose.Cells reflect updated theme colors. | Comparing cell foreground ARGB before and after Workbook.SetThemeColor to confirm style propagation. | Ensuring consistency of corporate branding colors across pivot tables after theme modifications.
// AI Prompts: Generate C# unit test code that asserts a pivot table cell color changes after modifying ThemeColorType.Accent1 with Aspose.Cells. | Explain the role of Worksheet.RefreshPivotTables in applying theme changes to existing pivot tables. | Provide a step‑by‑step guide to validate theme‑driven pivot table styles using Aspose.Cells for .NET.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsThemePivotValidation
{
    // This C# example creates a workbook, builds a pivot table with a built‑in style that uses theme accent colors, records the foreground ARGB of a data cell, changes the Accent1 theme color to red via Workbook.SetThemeColor, refreshes the pivot tables, re‑reads the cell color, and confirms the style automatically reflects the updated theme.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Drink");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("Snack");
            sheet.Cells["B4"].PutValue(50);

            // Add a pivot table
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D2", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Amount

            // Apply a built‑in style that uses theme accent colors
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium10;

            // Calculate the pivot data so the table is populated
            pivot.CalculateData();

            // Capture the foreground ARGB color of a data cell before theme change
            CellArea dataRange = pivot.DataBodyRange;
            int sampleRow = dataRange.StartRow;
            int sampleCol = dataRange.StartColumn;
            Style beforeStyle = sheet.Cells[sampleRow, sampleCol].GetStyle();
            int beforeColor = beforeStyle.ForegroundArgbColor;

            Console.WriteLine($"Foreground ARGB before theme change: 0x{beforeColor:X8}");

            // Change the theme accent color (Accent1) to Red
            workbook.SetThemeColor(ThemeColorType.Accent1, Color.Red);

            // Refresh pivot tables to ensure they reflect the new theme
            sheet.RefreshPivotTables();

            // Capture the foreground ARGB color of the same cell after theme change
            Style afterStyle = sheet.Cells[sampleRow, sampleCol].GetStyle();
            int afterColor = afterStyle.ForegroundArgbColor;

            Console.WriteLine($"Foreground ARGB after theme change: 0x{afterColor:X8}");

            // Validate that the color has changed (simple check)
            if (beforeColor != afterColor)
                Console.WriteLine("Pivot table style automatically adopted the new accent color.");
            else
                Console.WriteLine("Pivot table style did NOT adopt the new accent color.");

            // Save the workbook (lifecycle rule)
            workbook.Save("ThemePivotValidation.xlsx");
        }
    }
}
