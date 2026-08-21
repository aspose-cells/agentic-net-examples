// Title: C# Aspose.Cells: Preserve All Pivot Items by Using Max MissingItemsLimit
// Description: Shows how to build a workbook, insert sample rows, create a pivot table, assign fields, set the MissingItemsLimit property to the maximum enumeration value so the pivot cache keeps every possible entry after a refresh, then refreshes, calculates, and writes the file.
// Keywords: Aspose.Cells C# | PivotTable MissingItemsLimit | PivotMissingItemLimitType.Max | keep all pivot items | full item list refresh | Aspose.Cells pivot example
// Common Searches: Aspose.Cells missing items limit max | C# pivot table retain all categories after refresh | How to keep hidden rows in Aspose.Cells pivot | Set pivot cache to include all items
// Developer Intent: Configure a pivot table to retain the complete set of dimension values when the source data changes.
// Use Cases: Automatically display newly added categories without rebuilding the pivot. | Maintain rows with zero or missing values for consistent reporting layouts. | Create dashboards that always show the full set of dimension values regardless of data presence.
// AI Prompts: Write a C# code snippet that sets MissingItemsLimit to Max on an existing Aspose.Cells pivot table and updates it. | Describe how PivotMissingItemLimitType.Max influences pivot cache behavior and when to use it. | Show how to limit MissingItemsLimit to a specific number instead of Max in Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Pivot;

// Shows how to build a workbook, insert sample rows, create a pivot table, assign fields, set the MissingItemsLimit property to the maximum enumeration value so the pivot cache keeps every possible entry after a refresh, then refreshes, calculates, and writes the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].Value = "Category";
        sheet.Cells["B1"].Value = "Value";
        sheet.Cells["A2"].Value = "A";
        sheet.Cells["B2"].Value = 10;
        sheet.Cells["A3"].Value = "B";
        sheet.Cells["B3"].Value = 20;
        sheet.Cells["A4"].Value = "C";
        sheet.Cells["B4"].Value = 30;

        // Add a pivot table to the worksheet
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Configure the pivot table fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Value as data field

        // Set MissingItemsLimit to Max to retain all possible items during refresh
        pivotTable.MissingItemsLimit = PivotMissingItemLimitType.Max;

        // Refresh the pivot cache and calculate the pivot data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("PivotMissingItemsLimitDemo.xlsx");
    }
}
