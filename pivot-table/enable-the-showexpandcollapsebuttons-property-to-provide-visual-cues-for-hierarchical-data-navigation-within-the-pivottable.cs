using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExpandCollapseDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("SubCategory");
            sheet.Cells["C1"].PutValue("Amount");

            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue("Fruit");
            sheet.Cells["C2"].PutValue(120);

            sheet.Cells["A3"].PutValue("Food");
            sheet.Cells["B3"].PutValue("Vegetables");
            sheet.Cells["C3"].PutValue(80);

            sheet.Cells["A4"].PutValue("Beverage");
            sheet.Cells["B4"].PutValue("Tea");
            sheet.Cells["C4"].PutValue(50);

            sheet.Cells["A5"].PutValue("Beverage");
            sheet.Cells["B5"].PutValue("Coffee");
            sheet.Cells["C5"].PutValue(70);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure fields: Category as row, SubCategory as row, Amount as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "SubCategory");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Enable drilldown functionality (optional but often used together)
            pivotTable.EnableDrilldown = true;

            // Show expand/collapse buttons in the pivot table
            pivotTable.ShowDrill = true;

            // Refresh and calculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotTableShowExpandCollapse.xlsx");
        }
    }
}