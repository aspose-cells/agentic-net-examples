// Title: Reorder PivotTable Page Fields Programmatically with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, populates it with sample sales data, adds a PivotTable, inserts two page fields (Region and Category), moves the Category field to the first position using PageFields.Move, sets the page field layout and wrap count, adds row and data fields, refreshes the PivotTable, and saves the file with the updated filter priority.
// Keywords: Aspose.Cells | C# | PivotTable page field order | PageFields.Move | filter priority | .NET pivot table | page field wrap count | PrintOrderType.DownThenOver
// Common Searches: how to change page field order in Aspose.Cells pivot table | move pivot table filter to first position C# | set page field layout and wrap count Aspose.Cells | reorder pivot table page fields before adding to rows | Aspose.Cells PageFields.Move example
// Developer Intent: Programmatically adjust the order of page fields in a PivotTable to control filter priority before adding the same field to other areas.
// Use Cases: Place the most important filter (e.g., Category) at the top of the page field list in a sales dashboard. | Customize page field layout for printable reports with multiple filters. | Ensure correct filter sequence when the same field is used in row or column areas.
// AI Prompts: Generate C# code using Aspose.Cells to move a page field from index 1 to index 0, set PageFieldOrder to DownThenOver, and set PageFieldWrapCount to 2. | Explain the impact of PageFields.Move on filter priority in a PivotTable and show how to verify the new order after RefreshData. | Provide a step‑by‑step guide to reorder several page fields and configure their layout settings in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This example creates a workbook, populates it with sample sales data, adds a PivotTable, inserts two page fields (Region and Category), moves the Category field to the first position using PageFields.Move, sets the page field layout and wrap count, adds row and data fields, refreshes the PivotTable, and saves the file with the updated filter priority.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Region";
            cells["B1"].Value = "Category";
            cells["C1"].Value = "Sales";

            string[] regions = { "North", "South", "East", "West" };
            string[] categories = { "A", "B", "C" };
            int currentRow = 2;

            foreach (string region in regions)
            {
                foreach (string category in categories)
                {
                    cells[currentRow, 0].Value = region;
                    cells[currentRow, 1].Value = category;
                    cells[currentRow, 2].Value = (currentRow - 1) * 100; // sample sales value
                    currentRow++;
                }
            }

            // Add a pivot table based on the data range
            PivotTableCollection pivotTables = worksheet.PivotTables;
            int pivotIndex = pivotTables.Add($"A1:C{currentRow - 1}", "E5", "PivotTable1");
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Add page fields in the original order: Region then Category
            pivotTable.AddFieldToArea(PivotFieldType.Page, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Page, "Category");

            // Reorder page fields: move Category (index 1) to the first position (index 0)
            // This must be done before adding the same field to another area.
            pivotTable.PageFields.Move(1, 0);

            // Optionally adjust the layout order and wrap count for page fields
            pivotTable.PageFieldOrder = PrintOrderType.DownThenOver;
            pivotTable.PageFieldWrapCount = 2;

            // Add a row field and a data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook with the reordered page fields
            workbook.Save("ReorderedPageFields.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
