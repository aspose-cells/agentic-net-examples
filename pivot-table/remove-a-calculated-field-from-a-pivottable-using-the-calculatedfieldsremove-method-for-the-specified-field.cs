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
        sheet.Cells["B3"].PutValue(200);

        // Add a pivot table covering the data range and place it at D5
        int pivotIndex = sheet.PivotTables.Add("A1:B3", "D5", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add the base fields to the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Add a calculated field named "DoubleSales"
        string calculatedFieldName = "DoubleSales";
        pivotTable.AddCalculatedField(calculatedFieldName, "=Sales*2", true);

        // Remove the calculated field from the Data area
        // (CalculatedFields.Remove is not exposed; RemoveField achieves the same result)
        pivotTable.RemoveField(PivotFieldType.Data, calculatedFieldName);

        // Refresh and calculate the pivot table to reflect changes
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the updated pivot table
        workbook.Save("RemovedCalculatedField.xlsx");
    }
}