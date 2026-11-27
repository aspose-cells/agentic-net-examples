// Title: Toggle slicer item selection in Aspose.Cells C# based on database query results
// Common Searches: how to set specific slicer items as selected in Aspose.Cells C# after reading values from SQL | Aspose.Cells C# programmatically select slicer items using a HashSet of database values | update slicer cache items selection in a pivot table workbook with data from a database | C# code to synchronize Excel slicer selections with an external data source using Aspose.Cells | toggle Aspose.Cells slicer selections based on a dynamic list of values
// Tags: Aspose.Cells slicer cache item selection | C# programmatic slicer item toggle | Excel pivot table slicer integration Aspose.Cells | populate slicer from database values C# | set slicer item selected property Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerToggleDemo
{
    // The example creates a workbook with sample data, adds a pivot table and a linked slicer, simulates fetching a set of values (e.g., "Apple" and "Date") from a database, iterates through the slicer's cache items and sets each item's Selected flag according to the fetched list, and saves the result as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a workbook with sample data, a pivot table and a slicer.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Sample data: Fruit and Quantity
                cells["A1"].PutValue("Fruit");
                cells["B1"].PutValue("Quantity");
                string[] fruits = { "Apple", "Banana", "Cherry", "Date", "Elderberry" };
                int[] qty = { 10, 20, 15, 5, 12 };
                for (int i = 0; i < fruits.Length; i++)
                {
                    cells[i + 1, 0].PutValue(fruits[i]);   // Column A
                    cells[i + 1, 1].PutValue(qty[i]);     // Column B
                }

                // Create a pivot table based on the data range
                PivotTableCollection pivots = sheet.PivotTables;
                // Destination cell "D3" (row 3, column 4 in Excel notation)
                int pivotIdx = pivots.Add("=Sheet1!A1:B6", "D3", "PivotTable1");
                PivotTable pivot = pivots[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
                pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
                pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

                // Refresh and calculate the pivot table
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the pivot table for the "Fruit" field
                int slicerIdx = sheet.Slicers.Add(pivot, "Fruit", "G1");
                Slicer slicer = sheet.Slicers[slicerIdx];

                // -----------------------------------------------------------------
                // 2. Retrieve the list of values that should be selected from the DB.
                // -----------------------------------------------------------------
                // NOTE: Database access is omitted for portability; replace with actual DB code as needed.
                HashSet<string> valuesToSelect = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                try
                {
                    // Example placeholder: pretend we fetched "Apple" and "Date" from the database
                    valuesToSelect.Add("Apple");
                    valuesToSelect.Add("Date");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Database access failed: " + ex.Message);
                    // Continue with an empty selection set if DB access fails.
                }

                // -----------------------------------------------------------------
                // 3. Toggle slicer items based on the retrieved values.
                // -----------------------------------------------------------------
                try
                {
                    SlicerCache cache = slicer.SlicerCache;
                    SlicerCacheItemCollection items = cache.SlicerCacheItems;

                    for (int i = 0; i < items.Count; i++)
                    {
                        SlicerCacheItem item = items[i];
                        bool shouldSelect = valuesToSelect.Contains(item.Value);
                        item.Selected = shouldSelect;
                        Console.WriteLine($"Item \"{item.Value}\" -> Selected = {shouldSelect}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error while toggling slicer items: " + ex.Message);
                }

                // -----------------------------------------------------------------
                // 4. Save the workbook.
                // -----------------------------------------------------------------
                string outPath = "SlicerToggleResult.xlsx";
                try
                {
                    // Ensure the directory exists
                    string outDir = Path.GetDirectoryName(outPath);
                    if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                    {
                        Directory.CreateDirectory(outDir);
                    }

                    workbook.Save(outPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved to {outPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Failed to save workbook: " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unexpected error: " + ex.Message);
            }
        }
    }
}
