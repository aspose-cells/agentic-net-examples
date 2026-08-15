// Title: Refresh a PivotTable after Changing its ConsolidationFunction – Aspose.Cells for .NET (C#)
// Description: C# sample that creates a workbook, inserts sample data, builds a PivotTable, calculates it, switches the data field's ConsolidationFunction (e.g., Sum → Count), calls RefreshData to update the cache, recalculates with CalculateData, and saves the result.
// Keywords: Aspose.Cells | PivotTable RefreshData | CalculateData | ConsolidationFunction | change aggregation C# | .NET Excel automation | programmatic pivot refresh | update pivot cache | Sum to Count Aspose.Cells | Excel PivotTable API
// Common Searches: Aspose.Cells refresh pivot after changing consolidation function | C# change PivotTable aggregation type programmatically | PivotTable.RefreshData vs CalculateData example | How to update pivot cache in Aspose.Cells .NET | Change pivot data field from Sum to Count using Aspose.Cells
// Developer Intent: Apply a new ConsolidationFunction to a PivotTable and refresh the cache so the aggregated values reflect the change.
// Use Cases: Generate a financial report, switch the amount field from Sum to Count, and output the updated totals. | Build an interactive dashboard where users select an aggregation (Sum, Count, Average) and the code updates the pivot accordingly. | Process a workbook with multiple pivots, programmatically set each data field to a different function, and ensure all tables display the correct calculations.
// AI Prompts: Write C# code using Aspose.Cells that changes a PivotTable data field to Average, refreshes the cache, and saves the workbook. | Explain when to call PivotTable.RefreshData() versus PivotTable.CalculateData() in Aspose.Cells. | Provide step‑by‑step instructions to modify a PivotTable's ConsolidationFunction and verify the changes in the saved Excel file.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// C# sample that creates a workbook, inserts sample data, builds a PivotTable, calculates it, switches the data field's ConsolidationFunction (e.g., Sum → Count), calls RefreshData to update the cache, recalculates with CalculateData, and saves the result.
class RefreshPivotAfterConsolidationChange
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Amount");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["A4"].PutValue("A");
            worksheet.Cells["B4"].PutValue(30);

            // Add a pivot table based on the data range
            int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Add a row field (Category) and a data field (Amount)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

            // Initial calculation to populate the pivot table
            pivotTable.CalculateData();

            // Change the consolidation function of the data field (e.g., from Sum to Count)
            PivotField dataField = pivotTable.DataFields[0];
            dataField.Function = ConsolidationFunction.Count;

            // Refresh the pivot cache and recalculate to reflect the new aggregation
            pivotTable.RefreshData();          // Correct method to refresh cache
            pivotTable.CalculateData();

            // Save the workbook with the updated pivot table
            workbook.Save("PivotConsolidationRefresh.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
