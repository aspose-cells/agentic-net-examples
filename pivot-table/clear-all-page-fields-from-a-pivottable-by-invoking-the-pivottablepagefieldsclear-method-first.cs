using Aspose.Cells;
using Aspose.Cells.Pivot;

class ClearPivotTablePageFields
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Region");
        sheet.Cells["B1"].PutValue("Product");
        sheet.Cells["C1"].PutValue("Sales");
        sheet.Cells["A2"].PutValue("North");
        sheet.Cells["B2"].PutValue("Car");
        sheet.Cells["C2"].PutValue(1000);
        sheet.Cells["A3"].PutValue("South");
        sheet.Cells["B3"].PutValue("Bike");
        sheet.Cells["C3"].PutValue(800);

        // Add a pivot table that includes a page field (Region)
        int ptIndex = sheet.PivotTables.Add("A1:C3", "E5", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[ptIndex];

        // Add fields to the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Page, "Region");   // page field
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");   // row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");    // data field

        // Build the pivot table
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Clear all page fields from the pivot table
        pivotTable.PageFields.Clear();

        // Recalculate after clearing page fields (optional but ensures consistency)
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("ClearPageFieldsDemo.xlsx");
    }
}