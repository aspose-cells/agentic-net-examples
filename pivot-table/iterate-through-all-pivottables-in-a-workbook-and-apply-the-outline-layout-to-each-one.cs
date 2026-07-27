using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotOutlineDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the collection of pivot tables on the current worksheet
                PivotTableCollection pivots = sheet.PivotTables;

                // Apply outline layout to each pivot table in the collection
                for (int i = 0; i < pivots.Count; i++)
                {
                    PivotTable pivot = pivots[i];
                    pivot.ShowInOutlineForm();   // Layout the pivot table in outline form
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}