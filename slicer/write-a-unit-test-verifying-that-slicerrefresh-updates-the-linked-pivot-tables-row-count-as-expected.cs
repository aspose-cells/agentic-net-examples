// Title: C# unit test with Aspose.Cells to confirm that Slicer.Refresh updates the linked pivot table row count
// Common Searches: how to write a C# unit test that checks slicer refresh changes pivot table rows in Aspose.Cells | Aspose.Cells C# example verifying slicer cache deselection updates pivot row count | unit testing workbook.RefreshAll effect on slicer‑linked pivot tables using Aspose.Cells | C# code to assert pivot table row count after selecting items in an Aspose.Cells slicer | testing Aspose.Cells slicer selection impact on pivot data body range in .NET
// Tags: Aspose.Cells slicer refresh pivot table row count | C# unit test slicer cache selection impact | Aspose.Cells workbook.RefreshAll linked pivot validation | pivot table data body range assertion Aspose.Cells | C# Aspose.Cells slicer cache item deselection | testing slicer‑pivot synchronization Aspose.Cells | Aspose.Cells pivot row count verification

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerRefreshDemo
{
    // Simple assertion helper to replace NUnit assertions.
    // Creates a workbook with fruit data, adds a pivot table summarizing amounts, links a slicer to the 'Fruit' field, deselects the 'Date' item via the slicer cache, refreshes the workbook, and asserts that the pivot table row count changes from four to three.
    public static class Assert
    {
        public static void AreEqual(int expected, int actual, string message = "")
        {
            if (expected != actual)
                throw new Exception($"Assert.AreEqual Failed. Expected: {expected}, Actual: {actual}. {message}");
        }

        public static void IsTrue(bool condition, string message = "")
        {
            if (!condition)
                throw new Exception($"Assert.IsTrue Failed. {message}");
        }
    }

    public class SlicerRefreshTest
    {
        public void Run()
        {
            try
            {
                // 1. Create a new workbook and populate sample data.
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Header
                cells["A1"].PutValue("Fruit");
                cells["B1"].PutValue("Amount");

                // Data rows (4 distinct fruits, each appears twice)
                string[] fruits = { "Apple", "Banana", "Cherry", "Date" };
                int row = 1;
                foreach (string fruit in fruits)
                {
                    // First occurrence
                    cells[row, 0].PutValue(fruit);
                    cells[row, 1].PutValue(10);
                    row++;

                    // Second occurrence
                    cells[row, 0].PutValue(fruit);
                    cells[row, 1].PutValue(20);
                    row++;
                }

                // 2. Add a pivot table that uses "Fruit" as a row field and sums "Amount".
                PivotTableCollection pivots = sheet.PivotTables;
                int pivotIndex = pivots.Add("=Sheet1!A1:B9", "D3", "FruitPivot");
                PivotTable pivot = pivots[pivotIndex];
                pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                // Default aggregation is Sum; explicit setting is unnecessary.

                // Refresh and calculate the pivot so it reflects the data.
                pivot.RefreshData();
                pivot.CalculateData();

                // 3. Verify initial row count (should be 4 distinct fruits).
                CellArea dataArea = pivot.DataBodyRange;
                int initialRowCount = dataArea.EndRow - dataArea.StartRow + 1;
                Assert.AreEqual(4, initialRowCount, "Initial pivot row count should match distinct fruit count.");

                // 4. Add a slicer linked to the pivot table for the "Fruit" field.
                int slicerIndex = sheet.Slicers.Add(pivot, "Fruit", "FruitSlicer");
                Slicer slicer = sheet.Slicers[slicerIndex];
                SlicerCache slicerCache = slicer.SlicerCache;

                // 5. Deselect one fruit (e.g., "Date") via the slicer cache.
                bool found = false;
                foreach (SlicerCacheItem item in slicerCache.SlicerCacheItems)
                {
                    if (item.Value.Equals("Date", StringComparison.OrdinalIgnoreCase))
                    {
                        item.Selected = false;
                        found = true;
                        break;
                    }
                }
                Assert.IsTrue(found, "Slicer cache item for 'Date' should be found.");

                // 6. Refresh the workbook (which refreshes slicer connections and pivot tables).
                workbook.RefreshAll();

                // 7. After refresh, the pivot should now have one less row (3 rows).
                dataArea = pivot.DataBodyRange;
                int refreshedRowCount = dataArea.EndRow - dataArea.StartRow + 1;
                Assert.AreEqual(3, refreshedRowCount, "Pivot row count after slicer refresh should reflect deselected item.");

                Console.WriteLine("Slicer refresh test passed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Slicer refresh test failed: {ex.Message}");
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            new SlicerRefreshTest().Run();
        }
    }
}
