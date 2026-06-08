using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class ValidatePivotTheme
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Amount");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(200);
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B4"].PutValue(300);

        // Add a pivot table at D2 and configure its fields
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D2", "MyPivot");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Amount as data field

        // Apply a built‑in style that uses theme accent colors
        pivotTable.PivotTableStyleType = PivotTableStyleType.PivotTableStyleLight1;

        // Populate the pivot table with calculated data
        pivotTable.CalculateData();

        // Save the workbook before changing the theme (optional verification step)
        workbook.Save("InitialPivot.xlsx");

        // Capture the original theme color applied to a pivot header cell (D2 is the top‑left cell of the pivot)
        Style originalHeaderStyle = worksheet.Cells["D2"].GetStyle();
        ThemeColor originalThemeColor = originalHeaderStyle.Font.ThemeColor;

        // Change the workbook theme's Accent1 color to a distinct value (Red)
        workbook.SetThemeColor(ThemeColorType.Accent1, Color.Red);

        // Refresh pivot tables so they re‑evaluate the theme colors
        worksheet.RefreshPivotTables();

        // Retrieve the header cell style after the theme change
        Style updatedHeaderStyle = worksheet.Cells["D2"].GetStyle();
        ThemeColor updatedThemeColor = updatedHeaderStyle.Font.ThemeColor;

        // Validate that the style still references Accent1 (theme type unchanged)
        bool typeUnchanged = updatedThemeColor != null && updatedThemeColor.ColorType == ThemeColorType.Accent1;

        // Validate that the actual theme color has been updated (the ThemeColor object reflects the new color)
        bool colorChanged = updatedThemeColor != null && !updatedThemeColor.Equals(originalThemeColor);

        Console.WriteLine($"Theme color type unchanged (should be Accent1): {typeUnchanged}");
        Console.WriteLine($"Theme color updated after theme change: {colorChanged}");

        // Save the final workbook to demonstrate the effect
        workbook.Save("PivotAfterThemeChange.xlsx");
    }
}