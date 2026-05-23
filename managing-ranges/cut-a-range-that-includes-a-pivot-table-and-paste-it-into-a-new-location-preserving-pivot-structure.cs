using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data that will be used for the pivot table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Amount");
        worksheet.Cells["A2"].PutValue("Fruits");
        worksheet.Cells["B2"].PutValue(1500);
        worksheet.Cells["A3"].PutValue("Vegetables");
        worksheet.Cells["B3"].PutValue(2500);

        // Add a pivot table at cell D1 (top‑left corner of the report)
        int pivotIndex = worksheet.PivotTables.Add("A1:B3", "D1", "SalesPivot");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Configure the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
        pivotTable.CalculateData(); // Populate the pivot table with data

        // Move the entire pivot table to a new location.
        // Row and column indices are zero‑based. Here we move it to cell B5 (row 4, column 1).
        pivotTable.MoveTo(4, 1);

        // Refresh the worksheet to ensure the pivot table reflects its new position
        worksheet.RefreshPivotTables();

        // Save the workbook
        workbook.Save("PivotTableMoved.xlsx");
    }
}