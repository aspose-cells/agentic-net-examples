using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class DisassociateSlicerFromPivot
{
    static void Main()
    {
        // Load the existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Proceed only if the worksheet contains at least one PivotTable and one Slicer
            if (sheet.PivotTables.Count > 0 && sheet.Slicers.Count > 0)
            {
                // Get the first PivotTable on the sheet (adjust if you need a specific one)
                PivotTable pivot = sheet.PivotTables[0];

                // Remove the connection between each slicer and the pivot table
                foreach (Slicer slicer in sheet.Slicers)
                {
                    slicer.RemovePivotConnection(pivot);
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}