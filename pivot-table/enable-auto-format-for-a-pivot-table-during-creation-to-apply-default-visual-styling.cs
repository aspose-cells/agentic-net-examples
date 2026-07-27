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
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(200);
        worksheet.Cells["A4"].PutValue("A");
        worksheet.Cells["B4"].PutValue(150);

        // Add a pivot table to the worksheet
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Configure the pivot table fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

        // Enable automatic formatting (default visual style)
        pivotTable.IsAutoFormat = true;
        // Optionally specify a particular auto‑format type
        pivotTable.AutoFormatType = PivotTableAutoFormatType.Classic;

        // Populate the pivot table with calculated data
        pivotTable.CalculateData();

        // Save the workbook with the formatted pivot table
        workbook.Save("PivotTableAutoFormatDemo.xlsx");
    }
}