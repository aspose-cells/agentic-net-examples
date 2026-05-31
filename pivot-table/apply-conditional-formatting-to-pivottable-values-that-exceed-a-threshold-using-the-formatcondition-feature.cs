using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using System.Drawing;

namespace AsposeCellsPivotConditionalFormatting
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Product";
            cells["B1"].Value = "Region";
            cells["C1"].Value = "Sales";

            cells["A2"].Value = "Laptop";
            cells["B2"].Value = "North";
            cells["C2"].Value = 1500;

            cells["A3"].Value = "Laptop";
            cells["B3"].Value = "South";
            cells["C3"].Value = 800;

            cells["A4"].Value = "Phone";
            cells["B4"].Value = "North";
            cells["C4"].Value = 1200;

            cells["A5"].Value = "Phone";
            cells["B5"].Value = "South";
            cells["C5"].Value = 600;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot fields
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            pivot.AddFieldToArea(PivotFieldType.Column, "Region");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Add a conditional format to the pivot table
            int formatIdx = pivot.ConditionalFormats.Add();
            PivotConditionalFormat pcf = pivot.ConditionalFormats[formatIdx];

            // Apply the format to the data field (Sales)
            pcf.AddFieldArea(PivotFieldType.Data, pivot.DataFields[0]);

            // Define the condition: highlight cells with Sales >= 1000
            int conditionIdx = pcf.FormatConditions.AddCondition(FormatConditionType.CellValue);
            FormatCondition condition = pcf.FormatConditions[conditionIdx];
            condition.Operator = OperatorType.GreaterOrEqual;
            condition.Formula1 = "1000";

            // Set the style for the condition (red background)
            condition.Style.BackgroundColor = Color.Red;

            // Refresh and calculate the pivot table to apply formatting
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableConditionalFormatting.xlsx");
        }
    }
}