// Title: How to build a consolidated PivotTable from multiple worksheets with a Sum aggregation using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to create a PivotTable that consolidates data from two worksheets, sets the Amount field to use the Sum consolidation function, and saves the workbook. | Show how to add a PivotTable to a new worksheet in Aspose.Cells by providing multiple source ranges, assigning row and data fields, refreshing the pivot cache, and calculating the results.
// Common Searches: Aspose.Cells C# create pivot table from multiple sheets with sum function | Consolidate data from several worksheets into one pivot table using Aspose.Cells .NET | Set ConsolidationFunction.Sum for pivot data field in Aspose.Cells example | Add PivotTable with multiple source ranges in C# Aspose.Cells tutorial
// Tags: Aspose.Cells multi‑sheet pivot aggregation | C# set pivot data field ConsolidationFunction | Aspose.Cells add pivot with multiple source ranges | PivotTable refresh cache Aspose.Cells | Export workbook to .xlsx using Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook with two worksheets containing Category and Amount data, adds a third worksheet for the pivot, defines two source ranges, inserts a PivotTable that consolidates those ranges, adds a row field for Category and a data field for Amount, sets the data field's function to Sum, refreshes and calculates the pivot, and saves the workbook as ConsolidatedPivotTableDemo.xlsx.
    public class ConsolidatedPivotTableDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -------------------------------------------------
                // Worksheet 1 - Sample data
                // -------------------------------------------------
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";
                Cells cells1 = sheet1.Cells;
                cells1["A1"].PutValue("Category");
                cells1["B1"].PutValue("Amount");
                cells1["A2"].PutValue("A");
                cells1["B2"].PutValue(100);
                cells1["A3"].PutValue("B");
                cells1["B3"].PutValue(200);
                cells1["A4"].PutValue("C");
                cells1["B4"].PutValue(300);

                // -------------------------------------------------
                // Worksheet 2 - Sample data
                // -------------------------------------------------
                Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
                Cells cells2 = sheet2.Cells;
                cells2["A1"].PutValue("Category");
                cells2["B1"].PutValue("Amount");
                cells2["A2"].PutValue("A");
                cells2["B2"].PutValue(150);
                cells2["A3"].PutValue("B");
                cells2["B3"].PutValue(250);
                cells2["A4"].PutValue("D");
                cells2["B4"].PutValue(400);

                // -------------------------------------------------
                // Worksheet 3 - PivotTable destination
                // -------------------------------------------------
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotSheet");

                // Define multiple consolidation ranges
                string[] sourceRanges = {
                    "=Sheet1!A1:B4",
                    "=Sheet2!A1:B4"
                };

                // Create empty page fields collection (no page fields needed)
                PivotPageFields pageFields = new PivotPageFields();

                // Add a PivotTable that consolidates the above ranges
                int pivotIndex = pivotSheet.PivotTables.Add(
                    sourceRanges,          // multiple source ranges
                    false,                 // isAutoPage
                    pageFields,            // page fields (empty)
                    "A3",                  // destination cell
                    "ConsolidatedPivot"); // pivot table name

                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure the PivotTable fields
                // Row field: Category (field index 0)
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
                // Data field: Amount (field index 1)
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

                // Set the consolidation function for the data field (e.g., Sum)
                pivotTable.DataFields[0].Function = ConsolidationFunction.Sum;

                // Refresh pivot cache and calculate the pivot data (correct API)
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("ConsolidatedPivotTableDemo.xlsx");
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
            ConsolidatedPivotTableDemo.Run();
        }
    }
}
