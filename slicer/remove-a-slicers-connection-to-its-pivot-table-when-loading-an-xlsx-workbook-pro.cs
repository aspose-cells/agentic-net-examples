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

        // Access the worksheet that contains the slicer (adjust index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one slicer and one pivot table
        if (worksheet.Slicers.Count > 0 && worksheet.PivotTables.Count > 0)
        {
            // Get the first slicer
            Slicer slicer = worksheet.Slicers[0];

            // Get the first pivot table (the one the slicer is connected to)
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Remove the slicer's connection to the pivot table
            slicer.RemovePivotConnection(pivotTable);
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}