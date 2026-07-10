using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class CreatePivotTableDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Prepare source data worksheet
        Worksheet sourceSheet = workbook.Worksheets[0];
        sourceSheet.Name = "SourceData";

        // Fill sample data
        Cells srcCells = sourceSheet.Cells;
        srcCells["A1"].PutValue("Category");
        srcCells["B1"].PutValue("Amount");
        srcCells["A2"].PutValue("Food");
        srcCells["B2"].PutValue(100);
        srcCells["A3"].PutValue("Drink");
        srcCells["B3"].PutValue(150);
        srcCells["A4"].PutValue("Food");
        srcCells["B4"].PutValue(200);

        // Add a worksheet that will host the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

        // Build the source data reference (e.g., =SourceData!A1:B4)
        string sourceData = $"=SourceData!{sourceSheet.Cells.MaxDisplayRange.Address}";

        // Add the pivot table using the (sourceData, destCellName, tableName) overload
        PivotTableCollection pivots = pivotSheet.PivotTables;
        int pivotIndex = pivots.Add(sourceData, "A1", "MyPivotTable");

        // Configure the pivot table fields
        PivotTable pivot = pivots[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Refresh and calculate the pivot data
        pivotSheet.RefreshPivotTables();
        pivot.CalculateData();

        // Save the workbook
        workbook.Save("CreatedPivotTable.xlsx");
    }
}