using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class SlicerLinkedToPivot
{
    static void Main()
    {
        // Load an existing workbook that already contains a pivot table
        Workbook workbook = new Workbook("input.xlsx");

        // Assume the pivot table is on the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Retrieve the first pivot table in the worksheet
        if (sheet.PivotTables.Count == 0)
        {
            Console.WriteLine("No pivot tables found in the worksheet.");
            return;
        }
        PivotTable pivot = sheet.PivotTables[0];

        // Add a slicer linked to the pivot table.
        // The slicer will be placed with its upper‑left corner at cell E2
        // and will filter by the pivot field named "fruit".
        int slicerIndex = sheet.Slicers.Add(pivot, "E2", "fruit");
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Optional: set a caption and style for the slicer
        slicer.Caption = "Fruit Slicer";
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}