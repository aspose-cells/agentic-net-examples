// Title: Programmatically deselect a slicer item and confirm pivot table filter using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# how to deselect a slicer value and refresh linked pivot table | verify that a pivot table excludes a category after updating slicer selection with Aspose.Cells | C# code sample for checking pivot table rows after changing slicer filter in Aspose.Cells | remove Apple from slicer and ensure pivot table data updates using Aspose.Cells .NET | Aspose.Cells programmatic slicer manipulation and pivot table validation example
// Tags: Aspose.Cells slicer item deselection | Aspose.Cells pivot table refresh after slicer change | C# verify pivot table row removal | Aspose.Cells slicer cache manipulation | pivot table filter validation Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerPivotVerification
{
    // The example creates a workbook with sample data, adds a pivot table, links a slicer to the "Category" field, programmatically deselects the "Apple" slicer item, refreshes and recalculates the pivot table, then scans the pivot rows to confirm that the "Apple" category is no longer present before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and fill it with sample data
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Header
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Year");
                dataSheet.Cells["C1"].PutValue("Amount");

                // Sample rows
                string[,] rows = new string[,]
                {
                    { "Apple",  "2020", "50" },
                    { "Orange", "2020", "60" },
                    { "Apple",  "2021", "70" },
                    { "Banana", "2020", "80" },
                    { "Banana", "2021", "90" },
                    { "Orange", "2021", "100" }
                };

                for (int i = 0; i < rows.GetLength(0); i++)
                {
                    dataSheet.Cells[i + 1, 0].PutValue(rows[i, 0]); // Category
                    dataSheet.Cells[i + 1, 1].PutValue(rows[i, 1]); // Year
                    dataSheet.Cells[i + 1, 2].PutValue(rows[i, 2]); // Amount
                }

                // 2. Add a pivot table based on the data
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
                string sourceRange = "=Data!A1:C7";
                string pivotDestination = "A12";

                // Use the classic overload (still functional)
                int pivotIndex = pivotSheet.PivotTables.Add(sourceRange, pivotDestination, "CategoryPivot");
                PivotTable pivot = pivotSheet.PivotTables[pivotIndex];

                // Configure pivot fields
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Column, "Year");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

                // Refresh and calculate the pivot table
                pivot.RefreshData();               // Refreshes data from source to cache
                pivot.CalculateData();             // Calculates the pivot values

                // 3. Add a slicer linked to the pivot table for the "Category" field
                int slicerIndex = pivotSheet.Slicers.Add(pivot, "Category", "CategorySlicer");
                Slicer slicer = pivotSheet.Slicers[slicerIndex];

                // 4. Display initial slicer items state
                Console.WriteLine("Initial slicer items (all selected):");
                PrintSlicerItems(slicer);

                // 5. Update slicer items – deselect "Apple"
                SlicerCache cache = slicer.SlicerCache;
                SlicerCacheItemCollection items = cache.SlicerCacheItems;

                foreach (SlicerCacheItem item in items)
                {
                    if (item.Value.Equals("Apple", StringComparison.OrdinalIgnoreCase))
                    {
                        item.Selected = false; // Deselect Apple
                        Console.WriteLine($"Item \"{item.Value}\" deselected.");
                    }
                    else
                    {
                        item.Selected = true; // Ensure other items are selected
                    }
                }

                // 6. Refresh the pivot table so it reflects the new slicer filter
                pivot.RefreshData();
                pivot.CalculateData();

                // 7. Verify that the pivot table no longer contains data for "Apple"
                Console.WriteLine();
                Console.WriteLine("Verifying pivot table after slicer update...");

                // The pivot table starts at cell A12 (as defined above)
                Cell startCell = pivotSheet.Cells[pivotDestination];
                int startRow = startCell.Row;          // zero‑based row index
                int startColumn = startCell.Column;    // zero‑based column index

                bool appleRowFound = false;
                // Scan the first column of the pivot (Category axis) until an empty cell is hit
                for (int row = startRow + 1; ; row++)
                {
                    Cell categoryCell = pivotSheet.Cells[row, startColumn];
                    string category = categoryCell.StringValue?.Trim();

                    // Stop when we reach an empty row (pivot data finished)
                    if (string.IsNullOrEmpty(category))
                        break;

                    if (category.Equals("Apple", StringComparison.OrdinalIgnoreCase))
                    {
                        appleRowFound = true;
                        break;
                    }
                }

                if (appleRowFound)
                {
                    Console.WriteLine("Verification FAILED: \"Apple\" row is still present in the pivot table.");
                }
                else
                {
                    Console.WriteLine("Verification PASSED: \"Apple\" row is not present – slicer filter applied correctly.");
                }

                // 8. Save the workbook (optional – useful for visual inspection)
                string outPath = "SlicerPivotVerification.xlsx";
                workbook.Save(outPath);
                Console.WriteLine($"\nWorkbook saved to \"{outPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        private static void PrintSlicerItems(Slicer slicer)
        {
            SlicerCacheItemCollection items = slicer.SlicerCache.SlicerCacheItems;
            for (int i = 0; i < items.Count; i++)
            {
                SlicerCacheItem item = items[i];
                Console.WriteLine($"  [{i}] Value = \"{item.Value}\", Selected = {item.Selected}");
            }
        }
    }
}
