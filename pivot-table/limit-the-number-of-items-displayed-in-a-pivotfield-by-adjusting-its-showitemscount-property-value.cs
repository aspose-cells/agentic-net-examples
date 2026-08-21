// Title: Limit PivotField Items in Aspose.Cells (C#) – Show Top N Rows with AutoShow
// Description: Creates a workbook, adds sample data, builds a pivot table, and restricts the row field to a defined number of items by enabling AutoShow and setting AutoShowCount (or ShowItemsCount) before refreshing and saving the file.
// Keywords: Aspose.Cells | C# | PivotTable | PivotField | AutoShow | ShowItemsCount | limit pivot items | top N rows | Excel automation | pivot table filtering
// Common Searches: Aspose.Cells limit pivot field items | C# show top 3 rows in pivot table | PivotField ShowItemsCount example | How to use AutoShow in Aspose.Cells | Restrict pivot table rows Aspose.Cells
// Developer Intent: Display only a specific number of row items (e.g., top N) in a pivot table using Aspose.Cells.
// Use Cases: Financial report that shows only the top 5 product categories by revenue. | Dashboard view that hides low‑frequency categories to keep the pivot table concise. | User‑driven report where the viewer selects how many rows should be visible without rebuilding the pivot. | Export of large pivot tables with limited rows to improve performance and file size.
// AI Prompts: Write C# code with Aspose.Cells to limit a PivotField to the top 10 items using ShowItemsCount. | Explain the difference between IsAutoShow and ShowItemsCount for limiting pivot field items in Aspose.Cells. | Generate a method that updates AutoShowCount at runtime based on a user‑provided integer. | Provide a step‑by‑step guide to filter pivot rows to a specific count in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, builds a pivot table, and restricts the row field to a defined number of items by enabling AutoShow and setting AutoShowCount (or ShowItemsCount) before refreshing and saving the file.
    public class LimitPivotFieldItemsDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");
            cells["A2"].PutValue("A");
            cells["B2"].PutValue(100);
            cells["A3"].PutValue("B");
            cells["B3"].PutValue(200);
            cells["A4"].PutValue("C");
            cells["B4"].PutValue(300);
            cells["A5"].PutValue("D");
            cells["B5"].PutValue(400);
            cells["A6"].PutValue("E");
            cells["B6"].PutValue(500);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the "Category" field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add the "Amount" field to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Access the row field and limit displayed items
            PivotField rowField = pivotTable.RowFields[0];
            rowField.IsAutoShow = true;          // Enable auto‑show
            rowField.AutoShowCount = 3;          // Show only top 3 items
            rowField.AutoShowField = -1;         // Use the field itself for ranking
            rowField.IsAscendShow = true;        // Show top items (ascending)

            // Refresh the pivot cache and calculate the pivot table data
            pivotTable.RefreshData();            // Correct API to refresh data
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("LimitPivotFieldItemsDemo.xlsx");
        }
    }
}
