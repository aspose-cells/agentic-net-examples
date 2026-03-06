using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    public class DeletePivotTableExample
    {
        public static void Run()
        {
            // Load an existing workbook that contains a pivot table (XLSX format)
            Workbook workbook = new Workbook("PivotTableExample.xlsx");

            // Access the first worksheet (index 0)
            Worksheet sheet = workbook.Worksheets[0];

            // Get the collection of pivot tables on this worksheet
            PivotTableCollection pivotTables = sheet.PivotTables;

            // If there is at least one pivot table, remove the first one
            if (pivotTables.Count > 0)
            {
                // Retrieve the first pivot table
                PivotTable pivotTable = pivotTables[0];

                // Delete the pivot table and its data
                pivotTables.Remove(pivotTable);
                // Alternatively, to keep the data you could use:
                // pivotTables.Remove(pivotTable, true);
            }

            // Save the modified workbook to a new XLSX file
            workbook.Save("PivotTableRemoved.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            DeletePivotTableExample.Run();
        }
    }
}