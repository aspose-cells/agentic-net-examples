using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRemoval
{
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of pivot tables on the current worksheet
                PivotTableCollection pivots = sheet.PivotTables;

                // Loop backwards so that removal does not affect the index order
                for (int i = pivots.Count - 1; i >= 0; i--)
                {
                    PivotTable pt = pivots[i];

                    // Check if the pivot table name starts with the required prefix
                    if (pt.Name != null && pt.Name.StartsWith("Temp_"))
                    {
                        // Remove the pivot table and its data
                        pivots.Remove(pt);
                    }
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("OutputWorkbook.xlsx");
        }
    }
}