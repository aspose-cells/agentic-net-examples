// Title: C# – Build a Consolidated PivotTable from Multiple Worksheets Using Union Ranges (Aspose.Cells)
// Description: Creates a workbook with three worksheets, defines a union of their A1:B4 ranges, adds a pivot table on a fourth sheet, sets "Category" as a row field and "Value" as a summed data field, refreshes and calculates the pivot, then saves the file as ConsolidatedPivot.xlsx.
// Keywords: Aspose.Cells | C# PivotTable | union range | multiple worksheets | consolidated pivot | PivotTables.Add | Aspose.Cells example | Excel pivot from several sheets | .NET data consolidation | pivot table source ranges
// Common Searches: Aspose.Cells create pivot table from multiple sheets | C# union range pivot example | how to combine worksheets into one pivot using Aspose.Cells | PivotTables.Add with multiple source ranges .NET | consolidate data across sheets in Aspose.Cells
// Developer Intent: Create a pivot table that aggregates data from three separate worksheets by using a union range as the source.
// Use Cases: Summarize sales categories across regional worksheets in a single report. | Combine departmental expense data into a unified financial dashboard. | Merge product inventory lists from several sheets for a consolidated analysis.
// AI Prompts: Generate C# code with Aspose.Cells that adds a pivot table using a list of union ranges across multiple worksheets. | Explain how to add page fields to the union‑range pivot for interactive filtering. | Show how to change the data field aggregation from Sum to Average in the consolidated pivot table.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotUnionExample
{
    // Creates a workbook with three worksheets, defines a union of their A1:B4 ranges, adds a pivot table on a fourth sheet, sets "Category" as a row field and "Value" as a summed data field, refreshes and calculates the pivot, then saves the file as ConsolidatedPivot.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // -----------------------------
                // 1. Prepare three source sheets
                // -----------------------------
                // Sheet1
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";
                FillSourceData(sheet1, new[] { "A", "B", "C" }, new[] { 10, 20, 30 });

                // Sheet2
                Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
                FillSourceData(sheet2, new[] { "A", "B", "D" }, new[] { 15, 25, 35 });

                // Sheet3
                Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");
                FillSourceData(sheet3, new[] { "B", "C", "D" }, new[] { 12, 22, 32 });

                // ---------------------------------
                // 2. Create a worksheet for the pivot
                // ---------------------------------
                Worksheet pivotSheet = workbook.Worksheets.Add("ConsolidatedPivot");

                // -------------------------------------------------
                // 3. Define the union (multiple consolidation) ranges
                // -------------------------------------------------
                string[] sourceRanges = new string[]
                {
                    "=Sheet1!A1:B4",
                    "=Sheet2!A1:B4",
                    "=Sheet3!A1:B4"
                };

                // No auto page fields; create an empty PivotPageFields object
                PivotPageFields pageFields = new PivotPageFields();

                // Add the pivot table using the overload that accepts multiple ranges.
                // Destination cell A1 corresponds to row = 0, column = 0.
                int pivotIndex = pivotSheet.PivotTables.Add(
                    sourceRanges,          // sourceData (union ranges)
                    false,                // isAutoPage
                    pageFields,           // pageFields (empty)
                    0,                    // start row (A1)
                    0,                    // start column (A1)
                    "UnionPivotTable");   // pivot table name

                // -------------------------------------------------
                // 4. Configure the pivot table fields
                // -------------------------------------------------
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Add "Category" as Row field
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

                // Add "Value" as Data field (sum)
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

                // Refresh and calculate the pivot data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // -------------------------------------------------
                // 5. Save the workbook
                // -------------------------------------------------
                string outputPath = "ConsolidatedPivot.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }

        // Helper method to populate a worksheet with sample data
        private static void FillSourceData(Worksheet sheet, string[] categories, int[] values)
        {
            Cells cells = sheet.Cells;
            // Header
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Value");

            // Data rows
            for (int i = 0; i < categories.Length; i++)
            {
                cells[i + 1, 0].PutValue(categories[i]); // Column A (row index starts at 0)
                cells[i + 1, 1].PutValue(values[i]);    // Column B
            }
        }
    }
}
