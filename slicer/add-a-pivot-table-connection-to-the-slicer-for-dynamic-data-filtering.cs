using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

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
        cells["B1"].Value = "Amount";
        cells["A2"].Value = "A";
        cells["B2"].Value = 10;
        cells["A3"].Value = "B";
        cells["B3"].Value = 20;
        cells["A4"].Value = "A";
        cells["B4"].Value = 30;
        cells["A5"].Value = "B";
        cells["B5"].Value = 40;

        // Add a pivot table based on the data range
        PivotTableCollection pivots = sheet.PivotTables;
        int pivotIndex = pivots.Add("A1:B5", "D2", "MyPivotTable");
        PivotTable pivot = pivots[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer linked to the "Category" field of the pivot table
        int slicerIndex = sheet.Slicers.Add(pivot, "F2", "Category");
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Connect the slicer to the pivot table for dynamic filtering
        slicer.AddPivotConnection(pivot);

        // Save the workbook
        workbook.Save("PivotSlicerConnection.xlsx");
    }
}