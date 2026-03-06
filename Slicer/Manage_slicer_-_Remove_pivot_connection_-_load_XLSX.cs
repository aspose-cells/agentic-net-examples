using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class ManageSlicerRemovePivotConnection
{
    static void Main()
    {
        // Load the existing workbook (XLSX)
        Workbook workbook = new Workbook("input.xlsx");

        // Assume the slicer and its related pivot table are on the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Get the first slicer in the collection (if any)
        if (sheet.Slicers.Count == 0)
        {
            Console.WriteLine("No slicers found on the worksheet.");
            return;
        }
        Slicer slicer = sheet.Slicers[0];

        // Get the first pivot table in the worksheet (if any)
        if (sheet.PivotTables.Count == 0)
        {
            Console.WriteLine("No pivot tables found on the worksheet.");
            return;
        }
        PivotTable pivotTable = sheet.PivotTables[0];

        // Remove the connection between the slicer and the pivot table
        slicer.RemovePivotConnection(pivotTable);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}