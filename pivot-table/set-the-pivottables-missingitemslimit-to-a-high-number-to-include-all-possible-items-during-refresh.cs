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
        cells["A1"].Value = "Category";
        cells["B1"].Value = "Value";
        cells["A2"].Value = "A";
        cells["B2"].Value = 10;
        cells["A3"].Value = "B";
        cells["B3"].Value = 20;
        cells["A4"].Value = "C";
        cells["B4"].Value = 30;

        // Add a pivot table to the worksheet
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Configure the pivot table fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field

        // Set MissingItemsLimit to Max to retain all possible items during refresh
        pivotTable.MissingItemsLimit = PivotMissingItemLimitType.Max;

        // Refresh the pivot cache and calculate the pivot data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("PivotMissingItemsLimitDemo.xlsx");
    }
}