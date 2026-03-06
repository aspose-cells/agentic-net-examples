using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the worksheet that contains the pivot table (assumed first worksheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is at least one pivot table
        if (sheet.PivotTables.Count == 0)
        {
            Console.WriteLine("No pivot tables found in the worksheet.");
            return;
        }

        // Retrieve the first pivot table
        PivotTable pivot = sheet.PivotTables[0];

        // Add a slicer linked to the pivot table.
        // The slicer will be placed with its upper‑left corner at cell E2
        // and will filter the field named "Fruit" (replace with actual field name if different).
        int slicerIndex = sheet.Slicers.Add(pivot, "E2", "Fruit");
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Optional: customize slicer appearance
        slicer.Caption = "Fruit Slicer";
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

        // Refresh the slicer to reflect current pivot data
        slicer.Refresh();

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}