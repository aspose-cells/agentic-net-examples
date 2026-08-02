using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace HidePivotTableRibbonDemo
{
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
            sheet.Cells["A3"].PutValue("Beverage");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("Food");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["A5"].PutValue("Beverage");
            sheet.Cells["B5"].PutValue(70);

            // Add a new worksheet to host the pivot table
            int pivotSheetIndex = workbook.Worksheets.Add(SheetType.Worksheet);
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
            pivotSheet.Name = "PivotTable";

            // Create the pivot table
            int pivotIndex = pivotSheet.PivotTables.Add("=Sheet1!A1:B5", "C3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Hide the PivotTable field list (ribbon interface) for the whole workbook
            workbook.Settings.HidePivotFieldList = true;

            // Additionally, disable the field list for this specific pivot table
            pivotTable.EnableFieldList = false;

            // Save the workbook
            workbook.Save("HidePivotTableRibbonDemo.xlsx");
        }
    }
}