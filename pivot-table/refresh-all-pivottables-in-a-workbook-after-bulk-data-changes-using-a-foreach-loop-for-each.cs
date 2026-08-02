// Title: Refresh All PivotTables in Aspose.Cells (C#) After Bulk Data Changes
// Description: Shows how to programmatically refresh every PivotTable in an Aspose.Cells workbook using C# foreach loops after bulk updates to the source data, invoking RefreshData and CalculateData before saving.
// Keywords: Aspose.Cells | C# | RefreshData | CalculateData | PivotTable refresh | iterate pivot tables | bulk data update | workbook worksheets loop | .NET Excel automation
// Common Searches: Aspose.Cells refresh all pivot tables C# | How to update multiple PivotTables after data change .NET | foreach loop pivot tables Aspose.Cells | RefreshData CalculateData example | Refresh PivotTable programmatically Aspose
// Developer Intent: Programmatically refresh every PivotTable in a workbook after modifying its source data.
// Use Cases: Recalculate sales summary pivots after importing bulk sales figures. | Ensure financial report pivots reflect edited data across several worksheets before distribution. | Automate workbook preparation for export by looping through all sheets and refreshing each PivotTable.
// AI Prompts: Write C# code using Aspose.Cells to loop through all worksheets and refresh each PivotTable after changing cell values. | Show how to call RefreshData and CalculateData on every PivotTable in a workbook with Aspose.Cells. | Explain error handling when refreshing multiple PivotTables in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Shows how to programmatically refresh every PivotTable in an Aspose.Cells workbook using C# foreach loops after bulk updates to the source data, invoking RefreshData and CalculateData before saving.
class RefreshAllPivotTables
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet ws = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        ws.Cells["A1"].PutValue("Product");
        ws.Cells["B1"].PutValue("Sales");
        ws.Cells["A2"].PutValue("Apple");
        ws.Cells["B2"].PutValue(1000);
        ws.Cells["A3"].PutValue("Orange");
        ws.Cells["B3"].PutValue(2000);
        ws.Cells["A4"].PutValue("Banana");
        ws.Cells["B4"].PutValue(3000);

        // Add a pivot table to the worksheet
        int ptIndex = ws.PivotTables.Add("A1:B4", "E3", "PivotTable1");
        PivotTable pivot = ws.PivotTables[ptIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, 0);
        pivot.AddFieldToArea(PivotFieldType.Data, 1);

        // Simulate bulk data changes
        ws.Cells["B2"].PutValue(1500);
        ws.Cells["B3"].PutValue(2500);
        ws.Cells["B4"].PutValue(3500);

        // Refresh each pivot table using foreach loops
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            foreach (PivotTable pt in sheet.PivotTables)
            {
                // Refresh data from the source range and recalculate the pivot table
                pt.RefreshData();
                pt.CalculateData();
            }
        }

        // Save the updated workbook
        workbook.Save("RefreshedAllPivotTables.xlsx");
    }
}
