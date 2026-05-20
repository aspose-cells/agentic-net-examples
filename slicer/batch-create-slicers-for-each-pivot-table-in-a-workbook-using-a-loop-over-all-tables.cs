using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerBatch
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your source file)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of pivot tables on the current worksheet
                PivotTableCollection pivots = sheet.PivotTables;

                // Starting position for slicers on this sheet
                int slicerRow = 0;
                int slicerColumn = 0;

                // Loop through each pivot table
                foreach (PivotTable pivot in pivots)
                {
                    // Loop through each base field of the pivot table
                    foreach (PivotField baseField in pivot.BaseFields)
                    {
                        // Add a slicer for the current base field
                        // Using the overload: Add(PivotTable, int row, int column, PivotField baseField)
                        int slicerIndex = sheet.Slicers.Add(pivot, slicerRow, slicerColumn, baseField);
                        Slicer slicer = sheet.Slicers[slicerIndex];

                        // Optional: set a caption to identify the slicer
                        slicer.Caption = $"{pivot.Name}_{baseField.Name}_Slicer";

                        // Move to the next column for the next slicer
                        slicerColumn += 2; // adjust spacing as needed
                    }

                    // Reset column and move to next row block for the next pivot table
                    slicerColumn = 0;
                    slicerRow += 5; // adjust spacing as needed
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}