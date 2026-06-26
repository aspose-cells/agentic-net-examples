using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerBatchDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that already contains pivot tables
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each pivot table on the current worksheet
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    // Ensure the pivot table has at least one base field to create a slicer for
                    if (pivot.BaseFields.Count > 0)
                    {
                        // Choose a destination cell for the slicer.
                        // Here we place slicers sequentially in the first row, shifting columns by 2 per slicer.
                        int destRow = 0;
                        int destColumn = sheet.Slicers.Count * 2;

                        // Add a slicer using the first base field of the pivot table
                        // Overload: Add(PivotTable pivot, int row, int column, PivotField baseField)
                        int slicerIndex = sheet.Slicers.Add(pivot, destRow, destColumn, pivot.BaseFields[0]);

                        // Optional: customize the slicer (e.g., set a style)
                        Slicer slicer = sheet.Slicers[slicerIndex];
                        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
                    }
                }
            }

            // Save the modified workbook with the newly added slicers
            workbook.Save("output.xlsx");
        }
    }
}