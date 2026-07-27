using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class DeleteTempPivotTables
{
    static void Main()
    {
        // Load the workbook (replace with your source file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            PivotTableCollection pivotTables = sheet.PivotTables;

            // Iterate backwards so that removal does not affect the loop index
            for (int i = pivotTables.Count - 1; i >= 0; i--)
            {
                PivotTable pivot = pivotTables[i];

                // Check if the pivot table name starts with the "Temp_" prefix
                if (!string.IsNullOrEmpty(pivot.Name) && pivot.Name.StartsWith("Temp_"))
                {
                    // Remove the pivot table and its data
                    pivotTables.RemoveAt(i);
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}