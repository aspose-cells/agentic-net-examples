using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class DetachSlicerFromPivot
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Access the worksheet that contains the slicer (assumed first worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one slicer and one pivot table to work with
        if (worksheet.Slicers.Count > 0 && worksheet.PivotTables.Count > 0)
        {
            // Get the first slicer in the collection
            Slicer slicer = worksheet.Slicers[0];

            // Get the first pivot table that the slicer is currently connected to
            PivotTable pivotTable = worksheet.PivotTables[0];

            // Detach the slicer from the pivot table
            slicer.RemovePivotConnection(pivotTable);
        }

        // Save the workbook after detaching the slicer
        workbook.Save("OutputWorkbook.xlsx");
    }
}