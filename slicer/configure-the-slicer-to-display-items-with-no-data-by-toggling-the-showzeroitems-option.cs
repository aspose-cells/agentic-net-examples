using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class ConfigureSlicerShowZeroItems
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for a pivot table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["B4"].PutValue(30);

        // Add a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value field

        // Add a slicer linked to the pivot table for the "Category" field
        int slicerIndex = worksheet.Slicers.Add(pivotTable, "A1", "Category");
        Slicer slicer = worksheet.Slicers[slicerIndex];

        // Enable showing all items, even those without data
        slicer.ShowAllItems = true;

        // Configure how items with no data are displayed (e.g., natural order)
        slicer.ShowTypeOfItemsWithNoData = ItemsWithNoDataShowMode.Natural;

        // Save the workbook
        workbook.Save("SlicerShowZeroItems.xlsx");
    }
}