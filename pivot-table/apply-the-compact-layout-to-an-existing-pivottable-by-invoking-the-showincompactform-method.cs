using Aspose.Cells;
using Aspose.Cells.Pivot;

class ApplyCompactLayout
{
    static void Main()
    {
        // Load an existing workbook that already contains a pivot table
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet where the pivot table resides (here we use the first sheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is at least one pivot table
        if (sheet.PivotTables.Count > 0)
        {
            // Get the first pivot table
            PivotTable pivotTable = sheet.PivotTables[0];

            // Apply the compact layout to the pivot table
            pivotTable.ShowInCompactForm();

            // Refresh and recalculate so the layout change takes effect
            pivotTable.RefreshData();
            pivotTable.CalculateData();
        }

        // Save the workbook with the updated pivot table layout
        workbook.Save("output.xlsx");
    }
}