// Title: Refresh parent pivot table before child pivot table to keep totals accurate using Aspose.Cells for .NET (C#)
// AI Prompts: Update the source cells, then invoke parentPivot.RefreshData() and parentPivot.CalculateData() before calling childPivot.RefreshData() and childPivot.CalculateData() to synchronize totals. | After modifying data, programmatically recalculate a dependent child pivot by refreshing the parent pivot first with Aspose.Cells in C#. | Create two pivot tables that share the same range and ensure correct aggregation by refreshing and calculating the parent pivot prior to the child pivot.
// Common Searches: Aspose.Cells how to refresh parent pivot before child pivot in C# | C# calculate data for dependent pivot tables after source change using Aspose.Cells | sequence of RefreshData and CalculateData for multiple pivot tables Aspose.Cells .NET | ensure child pivot totals are correct after updating source data Aspose.Cells
// Tags: refresh order for dependent pivots Aspose.Cells | recompute totals for linked pivot after source change | C# linked pivot tables synchronization | Aspose.Cells multiple pivot tables operation sequence | accurate pivot totals after data modification

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotExample
{
    // The example creates a workbook with a parent and a child pivot table based on the same data range, modifies source values, refreshes and calculates the parent pivot first, then refreshes and calculates the child pivot to maintain correct totals, and finally saves the workbook as PivotParentChildRefresh.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot tables
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("Food");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Drink");
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["A4"].PutValue("Food");
                sheet.Cells["B4"].PutValue(150);
                sheet.Cells["A5"].PutValue("Drink");
                sheet.Cells["B5"].PutValue(70);

                // -------------------------------------------------
                // Create the parent pivot table
                // -------------------------------------------------
                int parentIndex = sheet.PivotTables.Add("A1:B5", "D1", "ParentPivot");
                PivotTable parentPivot = sheet.PivotTables[parentIndex];
                parentPivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category
                parentPivot.AddFieldToArea(PivotFieldType.Data, 1);  // Amount

                // -------------------------------------------------
                // Create a child pivot table that uses the same data source
                // -------------------------------------------------
                int childIndex = sheet.PivotTables.Add("A1:B5", "D10", "ChildPivot");
                PivotTable childPivot = sheet.PivotTables[childIndex];
                childPivot.AddFieldToArea(PivotFieldType.Row, 0);
                childPivot.AddFieldToArea(PivotFieldType.Data, 1);

                // Initial calculation so both pivots have data
                parentPivot.RefreshData();
                parentPivot.CalculateData();
                childPivot.RefreshData();
                childPivot.CalculateData();

                // -------------------------------------------------
                // Simulate a change in the source data
                // -------------------------------------------------
                sheet.Cells["B2"].PutValue(200); // Change Food amount for first row
                sheet.Cells["B4"].PutValue(250); // Change Food amount for second row

                // -------------------------------------------------
                // Refresh the parent pivot table first
                // -------------------------------------------------
                parentPivot.RefreshData();
                parentPivot.CalculateData();

                // -------------------------------------------------
                // Then refresh the child pivot table to ensure totals are accurate
                // -------------------------------------------------
                childPivot.RefreshData();
                childPivot.CalculateData();

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                workbook.Save("PivotParentChildRefresh.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
