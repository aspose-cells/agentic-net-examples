// Title: Create Multiple PivotTables Across Worksheets with Aspose.Cells for .NET (C#)
// Description: This example shows how to generate a new workbook, add several worksheets, populate each with sample data (A1:C5), and loop through the sheets to add a PivotTable using the `PivotTables.Add(sourceData, destinationCell, tableName)` overload. Each table is configured with row, column, and data fields and the workbook is saved as an XLSX file.
// Keywords: Aspose.Cells pivot table C# | add pivot table programmatically .NET | batch create pivot tables Aspose | loop worksheets pivot Aspose.Cells | PivotTables.Add overload example | automate pivot table generation | multiple worksheets pivot table
// Common Searches: how to add a pivot table to each worksheet using Aspose.Cells | batch generate pivot tables in a C# workbook | create identical pivot tables on multiple sheets programmatically | Aspose.Cells loop through worksheets to add pivot tables | C# example for adding multiple pivot tables
// Developer Intent: Programmatically add a separate PivotTable to every worksheet in a workbook by iterating over the sheets and invoking the `PivotTables.Add` method.
// Use Cases: Populate several sheets with raw data and instantly provide a summary view on each sheet. | Automate monthly reporting where the same pivot layout is required for different data sets. | Prepare a template workbook with pre‑configured PivotTables for downstream users or BI tools.
// AI Prompts: Generate code to apply a custom style to all PivotTables created in the loop. | Show how to refresh each PivotTable after modifying the source data programmatically. | Provide an example that adds a calculated field (e.g., Total = Amount * 1.2) to every PivotTable during creation.

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;

// This example shows how to generate a new workbook, add several worksheets, populate each with sample data (A1:C5), and loop through the sheets to add a PivotTable using the `PivotTables.Add(sourceData, destinationCell, tableName)` overload. Each table is configured with row, column, and data fields and the workbook is saved as an XLSX file.
class BatchPivotTables
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();

        // Loop through a set of worksheets and add a pivot table to each
        for (int wsIndex = 0; wsIndex < 3; wsIndex++)
        {
            Worksheet ws;
            if (wsIndex == 0)
            {
                // Use the default first worksheet
                ws = workbook.Worksheets[0];
                ws.Name = $"Sheet{wsIndex + 1}";
            }
            else
            {
                // Add additional worksheets
                ws = workbook.Worksheets.Add($"Sheet{wsIndex + 1}");
            }

            // Populate sample data (A1:C5) on the current worksheet
            Cells cells = ws.Cells;
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Item");
            cells["C1"].PutValue("Amount");

            for (int i = 2; i <= 5; i++)
            {
                cells[$"A{i}"].PutValue($"Cat{(i % 2) + 1}");
                cells[$"B{i}"].PutValue($"Item{i}");
                cells[$"C{i}"].PutValue(i * 10);
            }

            // Define the source data range string for the pivot table
            string sourceData = $"=Sheet{wsIndex + 1}!A1:C5";

            // Destination cell where the pivot table will be placed
            string destCell = "E3";

            // Unique pivot table name per worksheet
            string tableName = $"Pivot_{wsIndex + 1}";

            // Add a new pivot table using the Add(string, string, string) overload
            int pivotIndex = ws.PivotTables.Add(sourceData, destCell, tableName);

            // Retrieve the newly created pivot table and configure its fields
            PivotTable pivot = ws.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Column, "Item");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
        }

        // Save the workbook (lifecycle rule: save)
        workbook.Save("BatchPivotTables.xlsx");
    }
}
