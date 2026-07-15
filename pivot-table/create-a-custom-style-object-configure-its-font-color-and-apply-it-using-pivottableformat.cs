using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using System.Drawing;

namespace AsposeCellsPivotStyleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["A2"].Value = "Food";
            sheet.Cells["B2"].Value = 120;
            sheet.Cells["A3"].Value = "Drink";
            sheet.Cells["B3"].Value = 80;
            sheet.Cells["A4"].Value = "Food";
            sheet.Cells["B4"].Value = 150;
            sheet.Cells["A5"].Value = "Drink";
            sheet.Cells["B5"].Value = 70;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "MyPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Amount as data field

            // Create a custom style and set its font color
            Style customStyle = workbook.CreateStyle();
            customStyle.Font.Color = Color.DarkGreen; // Desired font color
            customStyle.Font.IsBold = true;           // Additional formatting (optional)

            // Define the pivot area to which the style will be applied
            PivotArea area = new PivotArea(pivotTable);
            area.RuleType = PivotAreaType.Normal;      // Apply to normal cells
            area.AxisType = PivotFieldType.Data;       // Target the data area (you can change as needed)

            // Apply the custom style to the specified pivot area
            pivotTable.Format(area, customStyle);

            // Save the workbook
            workbook.Save("PivotTableCustomStyle.xlsx");
        }
    }
}