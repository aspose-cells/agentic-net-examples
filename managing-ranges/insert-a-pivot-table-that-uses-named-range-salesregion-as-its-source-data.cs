using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook (empty)
        Workbook workbook = new Workbook();

        // The named range "SalesRegion" is assumed to be defined in the workbook.
        // If it is not defined, you can create it here, e.g.:
        // workbook.Worksheets[0].Cells.CreateRange("A1:C10").Name = "SalesRegion";

        // Add a worksheet that will contain the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("PivotSheet");

        // Define the source data as the named range
        string sourceData = "SalesRegion";

        // Upper‑left cell where the pivot table will be placed
        string destCell = "A1";

        // Name for the new pivot table
        string tableName = "SalesPivot";

        // Add the pivot table using the (string, string, string) overload
        int pivotIndex = pivotSheet.PivotTables.Add(sourceData, destCell, tableName);

        // Retrieve the created pivot table
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // (Optional) Configure fields – adjust field names to match those in the named range
        // pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
        // pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Refresh data and calculate the pivot table
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook to a file
        workbook.Save("PivotTableWithNamedRange.xlsx");
    }
}