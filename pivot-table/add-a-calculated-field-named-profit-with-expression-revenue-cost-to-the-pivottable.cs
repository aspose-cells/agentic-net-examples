// Title: Add a Profit Calculated Field to an Aspose.Cells PivotTable in C#
// Description: Learn how to create a workbook with Product, Revenue, and Cost data, generate a PivotTable, and add a calculated field named Profit that computes Revenue minus Cost. The example shows refreshing the cache, recalculating the PivotTable, and saving the result as an Excel file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# PivotTable | calculated field | Profit = Revenue - Cost | Excel PivotTable .NET | financial report automation | Aspose.Cells example | PivotTable refresh | PivotTable calculate data | add calculated field C#
// Common Searches: how to add a calculated field profit in Aspose.Cells PivotTable C# | Aspose.Cells profit = revenue - cost example | refresh pivot table after adding calculated field Aspose.Cells | C# code for PivotTable calculated field Aspose.Cells | Aspose.Cells PivotTable calculate data method
// Developer Intent: Implement a Profit column in a PivotTable by defining a calculated field that subtracts Cost from Revenue, then update the PivotTable to reflect the new data.
// Use Cases: Generate profit‑by‑product financial statements directly within Excel workbooks. | Create dynamic dashboards where profit margins are calculated on the fly without altering source tables. | Export analysis‑ready files for BI tools that require profit as a separate metric.
// AI Prompts: Write C# code with Aspose.Cells to add a calculated field Profit = [Revenue] - [Cost] to an existing PivotTable. | Explain how to refresh and recalculate a PivotTable after inserting a calculated field using Aspose.Cells for .NET. | Provide best‑practice error handling when adding calculated fields to a PivotTable with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Learn how to create a workbook with Product, Revenue, and Cost data, generate a PivotTable, and add a calculated field named Profit that computes Revenue minus Cost. The example shows refreshing the cache, recalculating the PivotTable, and saving the result as an Excel file using Aspose.Cells for .NET.
    public class AddProfitCalculatedField
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data with Revenue and Cost columns
                Cells cells = sheet.Cells;
                cells["A1"].PutValue("Product");
                cells["B1"].PutValue("Revenue");
                cells["C1"].PutValue("Cost");

                cells["A2"].PutValue("A");
                cells["B2"].PutValue(5000);
                cells["C2"].PutValue(3000);

                cells["A3"].PutValue("B");
                cells["B3"].PutValue(7000);
                cells["C3"].PutValue(4000);

                cells["A4"].PutValue("C");
                cells["B4"].PutValue(6000);
                cells["C4"].PutValue(3500);

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:C4", "E3", "PivotTable1");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Revenue");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Cost");

                // Add the calculated field "Profit" with the formula [Revenue] - [Cost]
                pivotTable.AddCalculatedField("Profit", "=Revenue-Cost", true);

                // Refresh and calculate the pivot table to reflect the new field
                pivotTable.RefreshData();      // Correct method to refresh the cache
                pivotTable.CalculateData();    // Recalculate pivot data

                // Save the workbook
                workbook.Save("PivotTable_With_Profit_CalculatedField.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AddProfitCalculatedField.Run();
        }
    }
}
