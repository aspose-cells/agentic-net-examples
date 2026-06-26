using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Region");
        worksheet.Cells["B1"].PutValue("Sales");
        worksheet.Cells["A2"].PutValue("North");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("South");
        worksheet.Cells["B3"].PutValue(150);
        worksheet.Cells["A4"].PutValue("East");
        worksheet.Cells["B4"].PutValue(200);

        // Add a pivot table based on the data
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "RegionPivot");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivotTable.CalculateData();

        // Add a slicer linked to the "Region" field of the pivot table
        int slicerIndex = worksheet.Slicers.Add(pivotTable, "F1", "Region");
        Slicer slicer = worksheet.Slicers[slicerIndex];

        // Change the slicer title to "Region Filter"
        slicer.Title = "Region Filter";

        // Save the updated workbook
        workbook.Save("SlicerTitleUpdated.xlsx");
    }
}