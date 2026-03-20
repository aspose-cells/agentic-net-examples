using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerRefreshDemo
{
    class Program
    {
        static void Main()
        {
            // Load the existing workbook that contains a slicer linked to a PivotTable
            Workbook workbook = new Workbook("input.xlsx");

            // Assume the slicer is on the first worksheet; adjust index if needed
            Worksheet sheet = workbook.Worksheets[0];

            // Example of updating source data that the slicer is based on
            // (Here we simply change a value in the data range)
            sheet.Cells["B2"].PutValue(999); // modify data that may affect the slicer

            // Access the slicer collection on the worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Refresh each slicer to reflect the updated data and recalculate its PivotTable
            for (int i = 0; i < slicers.Count; i++)
            {
                Slicer slicer = slicers[i];
                slicer.Refresh();
            }

            // Save the workbook with the refreshed slicer
            workbook.Save("output.xlsx");
        }
    }
}