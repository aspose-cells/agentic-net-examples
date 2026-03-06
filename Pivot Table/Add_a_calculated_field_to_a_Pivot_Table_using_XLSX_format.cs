using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["A4"].PutValue("Orange");
        sheet.Cells["B4"].PutValue(200);

        // Add a pivot table based on the data range A1:B4, place it at E3, and name it "PivotTable1"
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add the source fields to the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");   // Row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");    // Data field

        // Add a calculated field that doubles the Sales value and drag it to the data area
        pivotTable.AddCalculatedField("DoubleSales", "=Sales*2", true);

        // Refresh the pivot cache and calculate the pivot data so that the new field is populated
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook in XLSX format
        workbook.Save("PivotTableWithCalculatedField.xlsx");
    }
}