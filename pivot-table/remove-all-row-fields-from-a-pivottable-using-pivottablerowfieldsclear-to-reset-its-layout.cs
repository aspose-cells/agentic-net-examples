// Title: Clear All Row Fields in an Aspose.Cells PivotTable (C#) using PivotTable.RowFields.Clear
// Description: This C# example creates a workbook, adds sample data, builds a pivot table on A1:B4, inserts a row field (Category) and a data field (Amount), then removes every row field with PivotTable.RowFields.Clear(), refreshes and recalculates the pivot, and saves the file as PivotTableRowFieldsCleared.xlsx.
// Keywords: Aspose.Cells | C# | PivotTable.RowFields.Clear | remove row fields | reset pivot layout | pivot table manipulation | Excel automation | GitHub example | Aspose.Cells pivot table
// Common Searches: Aspose.Cells C# clear row fields from pivot table | PivotTable.RowFields.Clear usage example | how to reset pivot layout in Aspose.Cells | remove all row fields Aspose.Cells pivot | C# code to clear pivot table row fields
// Developer Intent: Programmatically delete every row field from a PivotTable to return it to a column‑only layout.
// Use Cases: Reset a pivot table after changing source data or field selections. | Prepare a workbook with a clean summary view before distribution. | Dynamically re‑configure pivot layouts for multiple reporting scenarios.
// AI Prompts: Write C# code that creates a pivot table with several row fields and then clears them using PivotTable.RowFields.Clear. | Explain how PivotTable.RowFields.Clear affects the pivot cache and what steps are needed to refresh and recalculate the pivot. | Show how to enumerate existing RowFields, log their captions, and then clear them in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    // This C# example creates a workbook, adds sample data, builds a pivot table on A1:B4, inserts a row field (Category) and a data field (Amount), then removes every row field with PivotTable.RowFields.Clear(), refreshes and recalculates the pivot, and saves the file as PivotTableRowFieldsCleared.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Fruit");
            sheet.Cells["A3"].PutValue("Fruit");
            sheet.Cells["A4"].PutValue("Vegetable");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(15);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add a row field and a data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // At this point the pivot table has a row field.
            // Remove all row fields to reset the layout
            pivotTable.RowFields.Clear();

            // Refresh and recalculate after clearing row fields
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTableRowFieldsCleared.xlsx");
        }
    }
}
