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

        // Assume the pivot table resides on the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is at least one pivot table
        if (sheet.PivotTables.Count == 0)
        {
            Console.WriteLine("No pivot tables found in the worksheet.");
            return;
        }

        // Get the first pivot table
        PivotTable pivot = sheet.PivotTables[0];

        // Add a slicer linked to the pivot table.
        // The slicer will be placed with its top‑left corner at cell E3
        // and will be based on the pivot field named "Fruit".
        int slicerIndex = sheet.Slicers.Add(pivot, "E3", "Fruit");
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Optional: customize slicer appearance
        slicer.Caption = "Fruit Slicer";
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

        // Save the workbook with the new slicer
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}