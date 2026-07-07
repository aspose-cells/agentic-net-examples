using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Amount");
        worksheet.Cells["A2"].PutValue("Fruits");
        worksheet.Cells["B2"].PutValue(1500);
        worksheet.Cells["A3"].PutValue("Vegetables");
        worksheet.Cells["B3"].PutValue(2500);

        // Add a pivot table starting at cell D1
        int pivotIndex = worksheet.PivotTables.Add("A1:B3", "D1", "SalesPivot");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
        pivotTable.CalculateData();

        // Move (cut) the pivot table to a new location, preserving its structure
        pivotTable.MoveTo("G5");          // Destination cell for the upper‑left corner
        pivotTable.CalculateData();      // Recalculate after moving

        // Save the workbook
        workbook.Save("PivotCutPasteDemo.xlsx");
    }
}