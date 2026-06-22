using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class RemoveCustomGroupings
{
    static void Main()
    {
        // Load the workbook containing the pivot table
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one pivot table
        if (worksheet.PivotTables.Count == 0)
        {
            Console.WriteLine("No pivot tables found in the worksheet.");
            return;
        }

        // Access the first pivot table
        PivotTable pivotTable = worksheet.PivotTables[0];

        // Helper action to ungroup a field if it has grouping settings
        Action<PivotField> ungroupIfGrouped = field =>
        {
            if (field.GroupSettings != null)
            {
                // Remove all custom groupings from the field
                field.Ungroup();
            }
        };

        // Ungroup all row fields
        foreach (PivotField field in pivotTable.RowFields)
        {
            ungroupIfGrouped(field);
        }

        // Ungroup all column fields
        foreach (PivotField field in pivotTable.ColumnFields)
        {
            ungroupIfGrouped(field);
        }

        // Ungroup any base fields that might be grouped
        foreach (PivotField field in pivotTable.BaseFields)
        {
            ungroupIfGrouped(field);
        }

        // Recalculate the pivot table data after ungrouping
        pivotTable.CalculateData();

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}