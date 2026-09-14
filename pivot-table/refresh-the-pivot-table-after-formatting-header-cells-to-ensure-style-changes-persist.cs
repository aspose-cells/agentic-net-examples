// Title: How to preserve custom header formatting when refreshing an Aspose.Cells pivot table in C#
// AI Prompts: Create a pivot table from a data range, apply a bold Arial style with a light‑gray background to the column header, enable PreserveFormatting, and refresh the pivot tables using Aspose.Cells for .NET. | Programmatically format a pivot table header cell, set PreserveFormatting = true, then call worksheet.RefreshPivotTables() to keep the style after recalculation in C#.
// Common Searches: Aspose.Cells C# keep pivot table header style after RefreshPivotTables | preserve pivot table formatting when updating data with Aspose.Cells | how to set PreserveFormatting for pivot tables in Aspose.Cells .NET | refresh pivot tables without losing custom header colors Aspose.Cells | apply custom style to pivot table header and maintain after refresh C#
// Tags: Aspose.Cells pivot table header styling | maintain header style after pivot refresh | C# Aspose.Cells RefreshPivotTables usage | pivot table PreserveFormatting property .NET | format pivot table column header programmatically

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using System.Drawing;

namespace AsposeCellsPivotRefreshDemo
{
    // The example creates a workbook, adds sample data, builds a pivot table, formats the Category header with a bold Arial font and light‑gray background, enables PreserveFormatting, refreshes all pivot tables in the worksheet, and saves the result as PivotTableHeaderFormatted.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample source data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Vegetable");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("Fruit");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["A5"].PutValue("Vegetable");
            sheet.Cells["B5"].PutValue(70);

            // Add a pivot table based on the source data
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table (row field and data field)
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Calculate data so the pivot table is populated
            pivotTable.CalculateData();

            // ------------------------------------------------------------
            // Format the header cell of the pivot table
            // Header cell coordinates are based on the pivot table's own grid.
            // Row index 1 (first data row after the pivot title) and column 0
            // correspond to the column header for the "Category" field.
            // ------------------------------------------------------------
            Style headerStyle = workbook.CreateStyle();
            headerStyle.Font.Name = "Arial";
            headerStyle.Font.Size = 12;
            headerStyle.Font.IsBold = true;
            headerStyle.ForegroundColor = Color.LightGray;
            headerStyle.Pattern = BackgroundType.Solid;

            // Apply the style to the header cell (row 1, column 0 in pivot coordinates)
            pivotTable.Format(1, 0, headerStyle);

            // Ensure that formatting is preserved when the pivot table is refreshed
            pivotTable.PreserveFormatting = true;

            // Refresh all pivot tables in the worksheet so that the style persists
            sheet.RefreshPivotTables();

            // Save the workbook
            workbook.Save("PivotTableHeaderFormatted.xlsx", SaveFormat.Xlsx);
        }
    }
}
