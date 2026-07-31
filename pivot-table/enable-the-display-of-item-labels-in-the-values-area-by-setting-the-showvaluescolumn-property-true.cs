// Title: Show row item labels in the values area of an Aspose.Cells PivotTable (C#) – Set ShowValuesColumn = true
// Description: C# example that creates a workbook, adds sample data, builds a PivotTable, and enables the ShowValuesColumn property so that item labels appear inside the values area before saving the file.
// Keywords: Aspose.Cells | C# PivotTable | ShowValuesColumn | item labels in values area | Excel pivot display | Aspose.Cells API | pivot table values column | set ShowValuesColumn true | Aspose.Cells tutorial | Excel automation C#
// Common Searches: Aspose.Cells set ShowValuesColumn true | display row labels in pivot values area C# | how to enable values column in Aspose pivot table | ShowValuesColumn property example | Aspose.Cells pivot table item labels
// Developer Intent: Enable the ShowValuesColumn property on a PivotTable so that row item names are shown alongside aggregated values in the values column.
// Use Cases: Financial reports that need item names next to totals in a single column. | Dashboards where row labels and their sums must appear together for quick scanning. | Cross‑tab layouts with multiple data fields where each value column should include its corresponding row label.
// AI Prompts: Generate C# code using Aspose.Cells to create a PivotTable and set ShowValuesColumn = true so row labels appear in the values area. | Explain the effect of ShowValuesColumn on PivotTable layout and how it interacts with ShowGrandTotals in Aspose.Cells. | Provide a step‑by‑step guide to format the values column header after enabling ShowValuesColumn in an Aspose.Cells workbook.

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;

// C# example that creates a workbook, adds sample data, builds a PivotTable, and enables the ShowValuesColumn property so that item labels appear inside the values area before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Item1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Item2");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("Item3");
            sheet.Cells["B4"].PutValue(30);

            // Add a pivot table covering the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Item as row field, Value as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Row field (Item)
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Data field (Value)

            // Save the workbook to a file
            workbook.Save("PivotShowValuesColumn.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
