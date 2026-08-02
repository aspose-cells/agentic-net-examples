// Title: Create a Consolidated PivotTable from Multiple Worksheets with ConsolidationFunction – Aspose.Cells for .NET (C#)
// Description: This C# example demonstrates how to build a PivotTable that pulls data from two worksheets, uses the multi‑range overload, assigns Category, Year, and Value fields, sets the data field's ConsolidationFunction to Sum, refreshes and calculates the pivot, and saves the workbook as ConsolidatedPivotTableDemo.xlsx.
// Keywords: Aspose.Cells PivotTable multiple worksheets | ConsolidationFunction Sum C# | Aspose.Cells create consolidated pivot | pivot table multi‑range source Aspose.Cells | C# .NET Excel pivot example
// Common Searches: Aspose.Cells pivot table from several sheets | set ConsolidationFunction for PivotTable in Aspose.Cells | consolidate data ranges into one pivot using Aspose.Cells | C# example of multi‑range PivotTable Aspose.Cells | how to refresh calculated pivot in Aspose.Cells
// Developer Intent: Generate a PivotTable that aggregates data from multiple worksheets and applies a sum consolidation function using Aspose.Cells for .NET.
// Use Cases: Combine regional sales worksheets into a single summary pivot. | Aggregate yearly financial figures from separate department sheets. | Build a dashboard that consolidates category‑year metrics across multiple data tables.
// AI Prompts: Show C# code to create a consolidated PivotTable with the ConsolidationFunction set to Average using Aspose.Cells. | How can I change the ConsolidationFunction of an existing Aspose.Cells PivotTable at runtime? | Explain the steps to refresh and recalculate a multi‑range PivotTable after updating source data in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // This C# example demonstrates how to build a PivotTable that pulls data from two worksheets, uses the multi‑range overload, assigns Category, Year, and Value fields, sets the data field's ConsolidationFunction to Sum, refreshes and calculates the pivot, and saves the workbook as ConsolidatedPivotTableDemo.xlsx.
    public class ConsolidatedPivotTableDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------
            // Prepare source worksheets
            // -------------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            FillData(sheet1, "A");

            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            FillData(sheet2, "B");

            // ---------------------------------
            // Create a worksheet for the pivot
            // ---------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotSheet");

            // Define the consolidation ranges (source data from both sheets)
            string[] sourceRanges = {
                "=Sheet1!A1:C5",
                "=Sheet2!A1:C5"
            };

            // No auto page fields – we will add them manually if needed
            PivotPageFields pageFields = new PivotPageFields();

            // Add the pivot table using the overload that accepts multiple ranges
            int pivotIndex = pivotSheet.PivotTables.Add(
                sourceRanges,          // multiple consolidation ranges
                false,                 // isAutoPage
                pageFields,            // page fields (empty in this case)
                "A3",                  // destination cell for the pivot table
                "ConsolidatedPivot"); // pivot table name

            // Get the created pivot table
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Year");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

            // Set the consolidation function for the data field (e.g., Sum)
            if (pivotTable.DataFields.Count > 0)
            {
                pivotTable.DataFields[0].Function = ConsolidationFunction.Sum;
            }

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("ConsolidatedPivotTableDemo.xlsx");
        }

        // Helper method to fill sample data into a worksheet
        private static void FillData(Worksheet sheet, string prefix)
        {
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Year");
            cells["C1"].PutValue("Value");

            // Sample rows
            for (int i = 2; i <= 5; i++)
            {
                cells[$"A{i}"].PutValue($"{prefix}Cat{i - 1}");
                cells[$"B{i}"].PutValue(2020 + (i - 2));
                cells[$"C{i}"].PutValue(i * 100);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                ConsolidatedPivotTableDemo.Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
