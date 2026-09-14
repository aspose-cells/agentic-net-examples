// Title: Refresh Aspose.Cells PivotTable after setting ConsolidationFunction to Average using C#
// AI Prompts: Generate C# code that sets a PivotTable data field's ConsolidationFunction to Average, then refreshes and recalculates the pivot using Aspose.Cells. | Show the sequence of Aspose.Cells API calls required to update a PivotTable after modifying its aggregation function in a .NET workbook. | Provide a complete C# example that creates a workbook, adds a pivot table, changes the data field to use the Average function, and invokes RefreshData and CalculateData.
// Common Searches: C# Aspose.Cells how to recalculate pivot table after changing consolidation function | RefreshData method usage for Aspose.Cells PivotTable after modifying data field function | Update pivot cache in Aspose.Cells when changing PivotField Function to Average | Aspose.Cells pivot table aggregation change to average and refresh
// Tags: Aspose.Cells PivotTable RefreshData usage | C# set PivotField ConsolidationFunction to Average | Aspose.Cells recalculate pivot after aggregation change | Refresh pivot cache Aspose.Cells .NET | PivotTable CalculateData after RefreshData

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook with sample data, adds a PivotTable, changes the data field's ConsolidationFunction to Average, then calls RefreshData followed by CalculateData to update the pivot cache before saving the file.
    class RefreshPivotAfterConsolidationChange
    {
        static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(30);

            // Add a pivot table based on the sample data
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Category as row field, Amount as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Row field: Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Data field: Amount

            // Initial calculation to populate the pivot table
            pivotTable.CalculateData();

            // Change the aggregation of the data field to Average
            PivotField amountField = pivotTable.DataFields[0];
            amountField.Function = ConsolidationFunction.Average; // set aggregation

            // Refresh the pivot cache and recalculate to reflect the new aggregation
            pivotTable.RefreshData();      // correct method to refresh pivot cache
            pivotTable.CalculateData();

            // Save the workbook with the updated pivot table
            string outputPath = "PivotRefreshAfterConsolidationChange.xlsx";
            workbook.Save(outputPath);
        }
    }
}
