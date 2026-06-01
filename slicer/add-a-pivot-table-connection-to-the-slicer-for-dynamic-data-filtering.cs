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
        cells["B1"].Value = "Product";
        cells["C1"].Value = "Sales";

        cells["A2"].Value = "Electronics";
        cells["B2"].Value = "Laptop";
        cells["C2"].Value = 1200;

        cells["A3"].Value = "Electronics";
        cells["B3"].Value = "Phone";
        cells["C3"].Value = 800;

        cells["A4"].Value = "Furniture";
        cells["B4"].Value = "Chair";
        cells["C4"].Value = 150;

        // Add a pivot table based on the data range
        PivotTableCollection pivotTables = sheet.PivotTables;
        int pivotIndex = pivotTables.Add("A1:C4", "E1", "SalesPivot");
        PivotTable pivot = pivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Row, "Product");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer for the "Category" field (placed at cell F1 -> row 0, column 5)
        int slicerIndex = sheet.Slicers.Add(pivot, 0, 5, "Category");
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Connect the slicer to the pivot table for dynamic filtering
        slicer.AddPivotConnection(pivot);

        // Save the workbook
        workbook.Save("PivotSlicerConnection.xlsx");
    }
}