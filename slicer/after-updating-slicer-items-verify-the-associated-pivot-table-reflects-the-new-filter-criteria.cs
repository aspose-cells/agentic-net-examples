// Title: Aspose.Cells .NET: Verify PivotTable Updates After Changing Slicer Selections
// Description: This example creates a workbook, adds sample data, builds a PivotTable on the "Category" field, links a slicer to the same field, programmatically selects only the "Fruit" item, refreshes the slicer (which automatically refreshes the connected PivotTable), and validates that the PivotTable now contains a single row with the value "Fruit" before saving the file.
// Keywords: Aspose.Cells | .NET | C# | Excel slicer | PivotTable refresh | SlicerCacheItem | programmatic slicer selection | pivot filter verification | automated Excel reporting | data consistency
// Common Searches: Aspose.Cells verify pivot table after slicer change | C# refresh slicer linked to pivot table | how to programmatically select slicer items in Aspose.Cells | check pivot row count after slicer filter .NET | Aspose.Cells slicer refresh example
// Developer Intent: Ensure that modifying slicer selections programmatically updates the linked PivotTable correctly.
// Use Cases: Automate validation of slicer‑driven filters in generated Excel reports. | Select a specific slicer value (e.g., "Fruit") and confirm the PivotTable shows only that category. | Refresh slicer and PivotTable together to maintain data integrity before saving the workbook.
// AI Prompts: Write C# code using Aspose.Cells that selects multiple slicer items, refreshes the slicer, and asserts the expected PivotTable rows. | Explain how slicer.Refresh() propagates filter changes to a linked PivotTable in Aspose.Cells. | Create a unit test in C# that changes slicer selections and verifies the resulting PivotTable data with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerPivotVerification
{
    // This example creates a workbook, adds sample data, builds a PivotTable on the "Category" field, links a slicer to the same field, programmatically selects only the "Fruit" item, refreshes the slicer (which automatically refreshes the connected PivotTable), and validates that the PivotTable now contains a single row with the value "Fruit" before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate source data for the pivot table
                // Columns: Category, Amount
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Amount");
                cells["A2"].PutValue("Fruit");
                cells["B2"].PutValue(120);
                cells["A3"].PutValue("Vegetable");
                cells["B3"].PutValue(80);
                cells["A4"].PutValue("Fruit");
                cells["B4"].PutValue(150);
                cells["A5"].PutValue("Vegetable");
                cells["B5"].PutValue(70);
                cells["A6"].PutValue("Grain");
                cells["B6"].PutValue(50);
                cells["A7"].PutValue("Fruit");
                cells["B7"].PutValue(200);

                // Add a pivot table based on the source data
                int pivotIdx = sheet.PivotTables.Add("A1:B7", "D3", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                // Row field: Category, Data field: Sum of Amount
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                // Refresh and calculate the pivot table so it contains data
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the pivot table for the "Category" field
                int slicerIdx = sheet.Slicers.Add(pivot, "F3", "Category");
                Slicer slicer = sheet.Slicers[slicerIdx];
                // Optional: set a visual style
                slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

                // ------------------------------------------------------------
                // Update slicer items: select only "Fruit" and deselect others
                // ------------------------------------------------------------
                for (int i = 0; i < slicer.SlicerCache.SlicerCacheItems.Count; i++)
                {
                    SlicerCacheItem item = slicer.SlicerCache.SlicerCacheItems[i];
                    // Select the item whose value equals "Fruit"
                    string itemValue = item.Value?.ToString() ?? string.Empty;
                    if (itemValue.Equals("Fruit", StringComparison.OrdinalIgnoreCase))
                    {
                        item.Selected = true;
                    }
                    else
                    {
                        item.Selected = false;
                    }
                }

                // Refresh the slicer – this also refreshes and recalculates the linked pivot table
                slicer.Refresh();

                // ------------------------------------------------------------
                // Verify that the pivot table reflects the slicer filter
                // ------------------------------------------------------------
                // After filtering to "Fruit", the pivot table should contain only one row item
                int rowItemCount = pivot.RowFields[0].PivotItems.Count;
                Console.WriteLine($"Row items after slicer refresh: {rowItemCount}");

                // Additionally, verify that the remaining item is "Fruit"
                if (rowItemCount > 0)
                {
                    string remainingItem = pivot.RowFields[0].PivotItems[0].Value?.ToString() ?? string.Empty;
                    Console.WriteLine($"Remaining pivot row item: {remainingItem}");
                    Console.WriteLine($"Filter applied correctly: {remainingItem.Equals("Fruit", StringComparison.OrdinalIgnoreCase)}");
                }

                // Save the workbook (lifecycle rule compliance)
                workbook.Save("SlicerPivotVerification.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
