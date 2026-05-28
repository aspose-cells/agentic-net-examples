using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data for the pivot table
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Amount");
        cells["A2"].PutValue("A");
        cells["B2"].PutValue(100);
        cells["A3"].PutValue("B");
        cells["B3"].PutValue(200);
        cells["A4"].PutValue("A");
        cells["B4"].PutValue(150);
        cells["A5"].PutValue("B");
        cells["B5"].PutValue(250);

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add the Category field to the row area
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

        // Add the Amount field to the data area
        int dataFieldPos = pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
        PivotField amountField = pivotTable.DataFields[dataFieldPos];

        // Change the consolidation function to Average
        amountField.Function = ConsolidationFunction.Average;

        // Refresh the pivot cache and calculate the pivot data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the updated pivot table
        workbook.Save("PivotAverageFunction.xlsx");
    }
}