using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].Value = "Category";
        worksheet.Cells["A2"].Value = "Fruit";
        worksheet.Cells["A3"].Value = "Fruit";
        worksheet.Cells["A4"].Value = "Vegetable";

        worksheet.Cells["B1"].Value = "Sales";
        worksheet.Cells["B2"].Value = 120;
        worksheet.Cells["B3"].Value = 150;
        worksheet.Cells["B4"].Value = 200;

        // Add a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Add a slicer linked to the "Category" field of the pivot table
        int slicerIndex = worksheet.Slicers.Add(pivotTable, "E1", "Category");
        Slicer slicer = worksheet.Slicers[slicerIndex];

        // Resize the slicer using the Shape's point‑based properties for layout consistency
        slicer.Shape.WidthPt = 150;   // Width in points
        slicer.Shape.HeightPt = 80;   // Height in points

        // Save the workbook with the resized slicer
        workbook.Save("ResizedSlicer.xlsx");
    }
}