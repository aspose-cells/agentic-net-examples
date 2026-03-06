using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Use the first worksheet as the source data for the pivot table
        Worksheet dataSheet = workbook.Worksheets[0];

        // Create a pivot table (if one does not already exist)
        // Source range: A1:B10, destination cell: D1, name: PivotTable1
        int pivotIndex = dataSheet.PivotTables.Add("A1:B10", "D1", "PivotTable1");
        PivotTable pivot = dataSheet.PivotTables[pivotIndex];

        // Add fields to the pivot table: first column as Row field, second column as Data field
        pivot.AddFieldToArea(PivotFieldType.Row, 0);
        pivot.AddFieldToArea(PivotFieldType.Data, 1);

        // Add a new worksheet that will host the slicer
        Worksheet slicerSheet = workbook.Worksheets.Add("SlicerSheet");

        // Add a slicer linked to the pivot table.
        // Upper‑left corner of the slicer will start at cell A1,
        // and it will be based on the first field (index 0) of the pivot table.
        int slicerIndex = slicerSheet.Slicers.Add(pivot, "A1", 0);
        Slicer slicer = slicerSheet.Slicers[slicerIndex];

        // Update slicer properties as needed
        slicer.Caption = "Product Slicer";
        slicer.NumberOfColumns = 2;          // display items in two columns
        slicer.WidthPixel = 250;             // set width
        slicer.HeightPixel = 150;            // set height

        // Refresh the slicer to ensure it reflects the current pivot data
        slicer.Refresh();

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}