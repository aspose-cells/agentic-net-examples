// Title: Refresh a Single PivotTable Programmatically with Aspose.Cells for .NET (C#)
// Description: Shows how to create a workbook, add source data, build a PivotTable, run an initial refresh, change the underlying cells, and then update only that PivotTable using RefreshData and CalculateData before saving the file.
// Keywords: Aspose.Cells C# PivotTable refresh | RefreshData method | CalculateData method | update pivot cache .NET | single pivot refresh programmatically | Aspose.Cells PivotTable API | C# workbook pivot update | Excel pivot table refresh Aspose
// Common Searches: how to refresh one pivot table in Aspose.Cells | RefreshData vs CalculateData Aspose.Cells | C# update pivot cache after data change | programmatically refresh specific pivot table .NET | Aspose.Cells refresh only selected pivot tables
// Developer Intent: Update a specific PivotTable to reflect modified source data without affecting other pivots.
// Use Cases: After importing new sales figures, call RefreshData and CalculateData on the affected PivotTable to keep the report current. | In workbooks containing multiple pivots, refresh only the targeted table to improve performance. | Automate pivot updates in a nightly data‑processing job that writes results to an Excel file.
// AI Prompts: Generate C# code that refreshes a single PivotTable after changing its source range using Aspose.Cells. | Explain the roles of RefreshData and CalculateData when updating a PivotTable in Aspose.Cells. | Provide a script to refresh selected PivotTables in a workbook while leaving others unchanged.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Shows how to create a workbook, add source data, build a PivotTable, run an initial refresh, change the underlying cells, and then update only that PivotTable using RefreshData and CalculateData before saving the file.
class RefreshSpecificPivotTable
{
    static void Main()
    {
        // Create a new workbook and obtain the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate source data for the pivot table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(300);

        // Add a pivot table that uses the source range A1:B4 and place it at D3
        int pivotIdx = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];

        // Configure the pivot table: Product as row field, Sales as data field
        pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Column index 0 -> Product
        pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Column index 1 -> Sales

        // Initial refresh and calculation so the pivot shows correct data
        pivot.RefreshData();
        pivot.CalculateData();

        // Modify the underlying source data
        sheet.Cells["B2"].PutValue(150); // Updated sales for product A
        sheet.Cells["B3"].PutValue(250); // Updated sales for product B

        // Refresh only this specific pivot table to reflect the changes
        pivot.RefreshData();   // Refreshes the pivot cache from the data source
        pivot.CalculateData(); // Recalculates the pivot report

        // Save the workbook with the refreshed pivot table
        workbook.Save("RefreshedPivotTable.xlsx");
    }
}
