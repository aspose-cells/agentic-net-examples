using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class RemovePivotFieldGroupings
{
    static void Main()
    {
        // Load the workbook (replace with your source file)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one pivot table
        if (worksheet.PivotTables.Count == 0)
        {
            Console.WriteLine("No pivot tables found in the worksheet.");
            return;
        }

        // Get the first pivot table
        PivotTable pivotTable = worksheet.PivotTables[0];

        // Ungroup all row fields (removes any custom grouping)
        foreach (PivotField rowField in pivotTable.RowFields)
        {
            rowField.Ungroup();
        }

        // Ungroup all column fields (removes any custom grouping)
        foreach (PivotField colField in pivotTable.ColumnFields)
        {
            colField.Ungroup();
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}