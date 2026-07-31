// Title: Refresh a PivotTable after source data changes with Aspose.Cells (C#)
// Description: Shows how to modify worksheet cells, call RefreshData and CalculateData to update an Aspose.Cells PivotTable, and save the refreshed workbook.
// Keywords: Aspose.Cells PivotTable refresh | C# RefreshData | Aspose.Cells CalculateData | update pivot cache programmatically | refresh pivot after data edit | Aspose.Cells example | pivot table recalc C#
// Common Searches: Aspose.Cells how to refresh pivot table after editing data | C# refresh pivot cache Aspose.Cells | RefreshData CalculateData Aspose.Cells example | programmatically update pivot table in .NET | Aspose.Cells pivot table recalculate
// Developer Intent: Update a PivotTable to reflect changes made to its source range.
// Use Cases: Recalculate sales totals after adjusting figures before exporting a report. | Automate monthly workbook generation where source data is refreshed and the pivot must show current aggregates. | Integrate pivot refresh into a data‑processing pipeline to guarantee that all summary tables are up‑to‑date.
// AI Prompts: Generate C# code that changes source cells and refreshes all PivotTables in a workbook using Aspose.Cells. | Explain the role of RefreshData and CalculateData when a PivotTable’s source range is modified. | Create a reusable method that accepts a Workbook, updates a range, and ensures every PivotTable on each sheet is refreshed and recalculated.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Shows how to modify worksheet cells, call RefreshData and CalculateData to update an Aspose.Cells PivotTable, and save the refreshed workbook.
class RefreshPivotDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate source data for the pivot table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["A4"].PutValue("Apple");
        sheet.Cells["B4"].PutValue(150);

        // Add a pivot table based on the source range
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Configure the pivot table (row field and data field)
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product column
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales column

        // Initial calculation to populate the pivot table
        pivotTable.CalculateData();

        // ----- Modify the source data -----
        sheet.Cells["B2"].PutValue(120); // Change Apple sales from 100 to 120
        sheet.Cells["B3"].PutValue(250); // Change Banana sales from 200 to 250

        // Refresh the pivot cache and recalculate the pivot table
        pivotTable.RefreshData();   // Gather updated data from the source range
        pivotTable.CalculateData(); // Recalculate the pivot results

        // Save the workbook with the refreshed pivot table
        workbook.Save("RefreshedPivot.xlsx");
    }
}
