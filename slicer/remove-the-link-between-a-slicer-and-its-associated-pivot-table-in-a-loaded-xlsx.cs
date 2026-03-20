using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class RemoveSlicerPivotLink
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Proceed only if the worksheet contains at least one pivot table and one slicer
            if (sheet.PivotTables.Count > 0 && sheet.Slicers.Count > 0)
            {
                // Assume the first pivot table is the one linked to the slicers
                PivotTable pivot = sheet.PivotTables[0];

                // Remove the connection between each slicer and the pivot table
                for (int i = 0; i < sheet.Slicers.Count; i++)
                {
                    Slicer slicer = sheet.Slicers[i];
                    slicer.RemovePivotConnection(pivot);
                }
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}