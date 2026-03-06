using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class IteratePivotTables
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the collection of pivot tables on the current worksheet
            PivotTableCollection pivotTables = sheet.PivotTables;

            // Loop over each pivot table using the indexer
            for (int i = 0; i < pivotTables.Count; i++)
            {
                // Access the pivot table at the current index
                PivotTable pivot = pivotTables[i];

                // Example operation: output pivot table details
                Console.WriteLine($"Worksheet: {sheet.Name}");
                Console.WriteLine($"Pivot Table Name: {pivot.Name}");
                Console.WriteLine($"Pivot Table Range: {pivot.TableRange1.StartRow},{pivot.TableRange1.StartColumn} to {pivot.TableRange1.EndRow},{pivot.TableRange1.EndColumn}");

                // Refresh and recalculate the pivot table (optional)
                pivot.RefreshData();
                pivot.CalculateData();
            }
        }

        // Save the workbook after any modifications
        workbook.Save("output.xlsx");
    }
}