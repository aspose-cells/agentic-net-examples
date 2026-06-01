using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotMoveField
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Region");
            sheet.Cells["C1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("Bike");
            sheet.Cells["B2"].PutValue("North");
            sheet.Cells["C2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("Bike");
            sheet.Cells["B3"].PutValue("South");
            sheet.Cells["C3"].PutValue(1500);

            sheet.Cells["A4"].PutValue("Car");
            sheet.Cells["B4"].PutValue("North");
            sheet.Cells["C4"].PutValue(2000);

            sheet.Cells["A5"].PutValue("Car");
            sheet.Cells["B5"].PutValue("South");
            sheet.Cells["C5"].PutValue(2500);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Initial layout: Product in rows, Region in columns, Sales as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate to populate the pivot table
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // ----- Move the "Region" field from Column area to Row area -----
            // Remove the field from the Column area
            pivotTable.RemoveField(PivotFieldType.Column, "Region");
            // Add the same field to the Row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");

            // Refresh and calculate again after the layout change
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook with the updated pivot layout
            workbook.Save("PivotFieldMoved.xlsx");
        }
    }
}