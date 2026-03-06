using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one pivot table
        if (sheet.PivotTables.Count == 0)
        {
            Console.WriteLine("No pivot tables found in the worksheet.");
            return;
        }

        // Retrieve the first pivot table
        PivotTable pivot = sheet.PivotTables[0];

        // Add a slicer for the first base field of the pivot table.
        // Parameters: pivot table, row index, column index, base field index.
        // Here we place the slicer starting at cell E5 (row 5, column 5) and use base field index 0.
        int slicerIndex = sheet.Slicers.Add(pivot, 5, 5, 0);

        // Access the newly created slicer to set optional properties
        Slicer slicer = sheet.Slicers[slicerIndex];
        slicer.Caption = "Sample Slicer";
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}