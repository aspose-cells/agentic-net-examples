using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotThemeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Sample data for the pivot table ----------
            // Header row
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            // Data rows
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("B");
            sheet.Cells["B5"].PutValue(40);

            // ---------- Create the pivot table ----------
            int pivotIdx = sheet.PivotTables.Add("A1:B5", "E3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];

            // Add fields: Category as row field, Value as data field
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Ensure the style is applied to row headers
            pivot.ShowPivotStyleRowHeader = true;

            // ---------- Create a style that uses the theme's Dark1 color ----------
            // In the theme, Dark1 corresponds to the Background1 color.
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Pattern = BackgroundType.Solid;
            headerStyle.BackgroundThemeColor = new ThemeColor(ThemeColorType.Background1, 0); // Dark1 (Background1) with no tint
            headerStyle.Font.Color = Color.White; // Optional: make the font readable

            // ---------- Apply the style to each row header cell ----------
            // Row header cells contain the distinct category values ("A", "B", etc.).
            // We'll locate them using Find and then format the entire row of the pivot table that holds the header.
            Cell found = sheet.Cells.Find("A", null);
            if (found != null)
            {
                // The row index returned is the worksheet row where the header appears in the pivot table.
                // Format that row (column 0 corresponds to the row header area in the pivot table).
                pivot.Format(found.Row, 0, headerStyle);
            }

            found = sheet.Cells.Find("B", null);
            if (found != null)
            {
                pivot.Format(found.Row, 0, headerStyle);
            }

            // ---------- Save the workbook ----------
            workbook.Save("PivotTableRowHeaderDark1Theme.xlsx");
        }
    }
}