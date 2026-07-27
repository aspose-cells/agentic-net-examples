using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotReportFilterCaption
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Region");
            sheet.Cells["C1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["B2"].PutValue("North");
            sheet.Cells["C2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("Fruit");
            sheet.Cells["B3"].PutValue("South");
            sheet.Cells["C3"].PutValue(800);

            sheet.Cells["A4"].PutValue("Vegetable");
            sheet.Cells["B4"].PutValue("North");
            sheet.Cells["C4"].PutValue(600);

            sheet.Cells["A5"].PutValue("Vegetable");
            sheet.Cells["B5"].PutValue("South");
            sheet.Cells["C5"].PutValue(900);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");   // Row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");    // Data field

            // Add a report filter (page field)
            pivotTable.AddFieldToArea(PivotFieldType.Page, "Region");

            // Customize the caption of the report filter.
            // The caption displayed for a report filter is taken from the PageField's Name.
            // Assign a new descriptive string to it.
            pivotTable.PageFields[0].Name = "Sales Region Filter";

            // Ensure the caption is visible (optional, but commonly required)
            pivotTable.ShowRowHeaderCaption = true;

            // Refresh and calculate the pivot data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableReportFilterCaption.xlsx");
        }
    }
}