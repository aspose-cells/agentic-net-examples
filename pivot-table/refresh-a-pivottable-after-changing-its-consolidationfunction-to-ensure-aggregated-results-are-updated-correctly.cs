// Title: Aspose.Cells C# – Refresh a PivotTable after changing its ConsolidationFunction
// Description: This example creates a workbook, adds sample data, builds a PivotTable, changes the data field's aggregation to Average by setting the Function property, then calls RefreshData() and CalculateData() to recalculate the pivot so the new ConsolidationFunction is applied before saving the file.
// Keywords: Aspose.Cells refresh pivot | PivotTable ConsolidationFunction C# | update pivot aggregation Aspose.Cells | RefreshData method Aspose.Cells | CalculateData Aspose.Cells | change pivot function to average | recalculate pivot after function change
// Common Searches: How to refresh a PivotTable in Aspose.Cells after changing the ConsolidationFunction | Aspose.Cells PivotTable recalculate after setting Function to Average | RefreshData vs CalculateData in Aspose.Cells PivotTable | C# change pivot data field aggregation Aspose.Cells | Update pivot cache after modifying function Aspose.Cells
// Developer Intent: Ensure that a PivotTable reflects a new ConsolidationFunction by refreshing its cache and recalculating the data.
// Use Cases: After setting a data field's Function to Average (or Max, Count, etc.), call RefreshData() then CalculateData() to update the displayed values. | When switching a pivot's aggregation from Sum to Count, reload the cache with RefreshData before recalculating. | In automated reporting, modify the consolidation function of a pivot and immediately save the workbook with the updated calculations.
// AI Prompts: Show me C# code to change a PivotTable data field’s ConsolidationFunction to Max and refresh it using Aspose.Cells. | Provide a snippet that updates multiple PivotTable data fields’ functions and ensures the changes appear in the saved workbook. | Explain the difference between RefreshData() and CalculateData() in Aspose.Cells PivotTable processing.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// This example creates a workbook, adds sample data, builds a PivotTable, changes the data field's aggregation to Average by setting the Function property, then calls RefreshData() and CalculateData() to recalculate the pivot so the new ConsolidationFunction is applied before saving the file.
class RefreshPivotAfterConsolidationChange
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample source data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(30);

            // Add a pivot table based on the source range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "Pivot1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot: row field = Category, data field = Amount (default Sum)
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Category
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 -> Amount

            // Initial calculation so the pivot shows data
            pivot.CalculateData();

            // Change the aggregation (ConsolidationFunction) of the data field to Average
            // For data fields the aggregation is set via the Function property
            pivot.DataFields[0].Function = ConsolidationFunction.Average;

            // Refresh the pivot cache and recalculate to apply the new function
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook with the updated pivot table
            workbook.Save("PivotRefreshAfterConsolidationChange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
