// Title: Batch create PivotTables on multiple worksheets with Aspose.Cells for .NET (C#)
// Description: C# example that loops through worksheets, adds a PivotTable to each using the Add(sourceRange, destCell, tableName) overload, sets row/column/data fields, applies a style, refreshes all tables, and saves the workbook.
// Keywords: Aspose.Cells PivotTable C# | batch create pivot tables .NET | loop worksheets add pivot table | Aspose.Cells Add method overload | refresh all pivot tables Aspose | PivotTableStyleType Medium9 | programmatic pivot table generation | Excel automation Aspose.Cells
// Common Searches: How to add a pivot table to every sheet using Aspose.Cells | C# loop to create multiple pivot tables in a workbook | Aspose.Cells Add(string sourceRange, string destCell, string tableName) example | Apply the same PivotTable style to several tables with Aspose | Refresh all pivot tables in Aspose.Cells workbook
// Developer Intent: Create and configure a PivotTable on each worksheet by iterating over sheets and invoking Aspose.Cells' Add method.
// Use Cases: Generate identical pivot reports across several worksheets after populating each sheet with data. | Apply a consistent PivotTableStyle (e.g., Medium9) to all tables for uniform appearance. | Refresh every pivot table in one call before saving the workbook to ensure up‑to‑date calculations.
// AI Prompts: Write C# code with Aspose.Cells that loops through N worksheets, adds a pivot table to each using a dynamic source range, and sets row, column, and data fields. | Show how to apply PivotTableStyleType.PivotTableStyleMedium9 to multiple pivot tables created in a batch operation. | Explain how to construct source range strings for each sheet and refresh all pivot tables in an Aspose.Cells workbook before saving.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// C# example that loops through worksheets, adds a PivotTable to each using the Add(sourceRange, destCell, tableName) overload, sets row/column/data fields, applies a style, refreshes all tables, and saves the workbook.
class BatchPivotTables
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Number of worksheets to process
        int sheetCount = 3;

        for (int i = 0; i < sheetCount; i++)
        {
            // Get or create worksheet
            Worksheet ws = i == 0 ? workbook.Worksheets[0] : workbook.Worksheets.Add($"Sheet{i + 1}");

            // Populate sample data (A1:C5)
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Item");
            ws.Cells["C1"].PutValue("Amount");

            for (int r = 2; r <= 5; r++)
            {
                ws.Cells[$"A{r}"].PutValue($"Cat{(r % 3) + 1}");
                ws.Cells[$"B{r}"].PutValue($"Item{r + i * 10}");
                ws.Cells[$"C{r}"].PutValue(r * 10 + i * 100);
            }

            // Define source data range string (e.g., =Sheet1!A1:C5)
            string sourceRange = $"=Sheet{i + 1}!A1:C5";

            // Destination cell for the pivot table
            string destCell = "E3";

            // Unique pivot table name
            string tableName = $"Pivot_{i + 1}";

            // Add a new pivot table using the Add(string, string, string) overload
            int pivotIndex = ws.PivotTables.Add(sourceRange, destCell, tableName);
            PivotTable pivot = ws.PivotTables[pivotIndex];

            // Configure the pivot table fields
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Column, "Item");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Optional: apply a style
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
        }

        // Refresh all pivot tables in the workbook
        workbook.Worksheets.RefreshPivotTables();

        // Save the workbook
        workbook.Save("BatchPivotTables.xlsx");
    }
}
