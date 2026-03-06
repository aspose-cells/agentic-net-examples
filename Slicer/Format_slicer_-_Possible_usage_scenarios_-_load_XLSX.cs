using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // -------------------------------------------------
        // Prepare data worksheet (assumes data already exists)
        // -------------------------------------------------
        Worksheet dataSheet = workbook.Worksheets[0];

        // -------------------------------------------------
        // Create or get a worksheet for the pivot table
        // -------------------------------------------------
        Worksheet pivotSheet = workbook.Worksheets.Count > 1
            ? workbook.Worksheets[1]
            : workbook.Worksheets.Add("PivotSheet");

        // Add a pivot table using the Add method (row, column, baseFieldIndex overload)
        // Adjust the source range ("A1:B10") as needed for your data
        int pivotIdx = pivotSheet.PivotTables.Add("A1:B10", "C3", "PivotTable1");
        PivotTable pivot = pivotSheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, 0);   // First column as row field
        pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Second column as data field

        // -------------------------------------------------
        // Create or get a worksheet for the slicer
        // -------------------------------------------------
        Worksheet slicerSheet = workbook.Worksheets.Count > 2
            ? workbook.Worksheets[2]
            : workbook.Worksheets.Add("SlicerSheet");

        // Add a slicer using the Add method (destCellName, baseFieldName overload)
        int slicerIdx = slicerSheet.Slicers.Add(pivot, "E2", "fruit");
        Slicer slicer = slicerSheet.Slicers[slicerIdx];

        // ------------------------------
        // Format the slicer (properties)
        // ------------------------------
        slicer.StyleType = SlicerStyleType.SlicerStyleDark2;   // Apply a built‑in dark style
        slicer.Caption = "Fruit Selector";                    // Set the caption text
        slicer.NumberOfColumns = 2;                           // Display items in two columns
        slicer.ColumnWidthPixel = 120;                        // Width of each column (pixels)
        slicer.RowHeightPixel = 30;                           // Height of each row (pixels)
        slicer.LockedPosition = false;                        // Allow user to move the slicer
        slicer.Placement = PlacementType.FreeFloating;        // Free‑floating placement
        slicer.Refresh();                                     // Refresh to apply changes

        // -------------------------------------------------
        // Save the modified workbook
        // -------------------------------------------------
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}