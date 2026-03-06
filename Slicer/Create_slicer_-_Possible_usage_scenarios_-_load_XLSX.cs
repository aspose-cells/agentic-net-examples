using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("Input.xlsx");

        // Access the first worksheet (adjust index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one pivot table
        if (sheet.PivotTables.Count == 0)
        {
            Console.WriteLine("No pivot tables found in the worksheet.");
            return;
        }

        // Retrieve the first pivot table
        PivotTable pivot = sheet.PivotTables[0];

        // Use the first base field of the pivot table as the slicer field
        string baseFieldName = pivot.BaseFields[0].Name;

        // Add a slicer linked to the pivot table.
        // The slicer will be placed with its upper‑left corner at cell E2.
        int slicerIndex = sheet.Slicers.Add(pivot, "E2", baseFieldName);

        // Obtain the slicer object to customize its appearance
        Slicer slicer = sheet.Slicers[slicerIndex];
        slicer.Caption = $"{baseFieldName} Slicer";
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

        // Save the workbook with the new slicer
        workbook.Save("Output.xlsx");
    }
}