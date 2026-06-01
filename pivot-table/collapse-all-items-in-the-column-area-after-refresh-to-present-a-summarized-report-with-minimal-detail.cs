using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class CollapseColumnAreaDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Region");
        worksheet.Cells["B1"].PutValue("Product");
        worksheet.Cells["C1"].PutValue("Sales");
        worksheet.Cells["A2"].PutValue("East");
        worksheet.Cells["B2"].PutValue("Apple");
        worksheet.Cells["C2"].PutValue(1200);
        worksheet.Cells["A3"].PutValue("East");
        worksheet.Cells["B3"].PutValue("Banana");
        worksheet.Cells["C3"].PutValue(800);
        worksheet.Cells["A4"].PutValue("West");
        worksheet.Cells["B4"].PutValue("Apple");
        worksheet.Cells["C4"].PutValue(1500);
        worksheet.Cells["A5"].PutValue("West");
        worksheet.Cells["B5"].PutValue("Banana");
        worksheet.Cells["C5"].PutValue(700);

        // Add a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Define fields: Product as column, Region as row, Sales as data
        pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Refresh and calculate the pivot table data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Collapse all items in the column area after refresh
        foreach (PivotField columnField in pivotTable.ColumnFields)
        {
            // HideDetail collapses the details of the field
            columnField.HideDetail(true);
        }

        // Recalculate if needed after collapsing
        pivotTable.CalculateData();

        // Save the workbook with the collapsed column view
        workbook.Save("CollapsedColumnPivot.xlsx");
    }
}