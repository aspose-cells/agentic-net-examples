// Title: Synchronize two slicers linked to the same pivot table in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: asp.net how to link two slicers to the same pivot table with Aspose.Cells | c# code to share slicer cache between multiple slicers in Excel | programmatically select a slicer item and have other slicers update Aspose.Cells | example of slicer synchronization in an Excel file using Aspose.Cells for .NET | set slicer selection in C# and refresh pivot table with Aspose.Cells
// Tags: slicer cache sharing Aspose.Cells | programmatic slicer item selection C# | multiple slicers synchronization .NET | pivot table slicer binding Aspose.Cells | Excel slicer synchronization example

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// The example creates a workbook with fruit quantity data, builds a pivot table, adds two slicers bound to the same 'Fruit' field (sharing a slicer cache), programmatically selects 'Apple' in the first slicer, refreshes the pivot table, and shows that the second slicer automatically reflects the same selection before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the output workbook
            const string outputPath = "SlicerSyncDemo.xlsx";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // 1. Populate sample data that will be used for the pivot table
            // ------------------------------------------------------------
            Cells cells = sheet.Cells;
            cells["A1"].PutValue("Fruit");
            cells["B1"].PutValue("Quantity");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("Orange");
            cells["B3"].PutValue(15);
            cells["A4"].PutValue("Banana");
            cells["B4"].PutValue(20);
            cells["A5"].PutValue("Apple");
            cells["B5"].PutValue(5);
            cells["A6"].PutValue("Orange");
            cells["B6"].PutValue(8);
            cells["A7"].PutValue("Banana");
            cells["B7"].PutValue(12);

            // ------------------------------------------------------------
            // 2. Create a pivot table based on the data range
            // ------------------------------------------------------------
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIdx = pivots.Add("=Sheet1!$A$1:$B$7", "E3", "FruitPivot");
            PivotTable pivot = pivots[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // ------------------------------------------------------------
            // 3. Add two slicers for the same field ("Fruit") and connect them
            //    to the same pivot table. Because they share the same slicer
            //    cache, a selection change in one slicer is reflected in the
            //    other automatically.
            // ------------------------------------------------------------
            // First slicer placed at F3
            int slicerIdx1 = sheet.Slicers.Add(pivot, "F3", "Fruit");
            Slicer slicer1 = sheet.Slicers[slicerIdx1];

            // Second slicer placed at I3
            int slicerIdx2 = sheet.Slicers.Add(pivot, "I3", "Fruit");
            Slicer slicer2 = sheet.Slicers[slicerIdx2];

            // ------------------------------------------------------------
            // 4. Demonstrate synchronization programmatically:
            //    - Select only "Apple" in the first slicer.
            //    - Refresh the pivot table.
            //    - Verify that the second slicer shows the same selection.
            // ------------------------------------------------------------
            void SetSlicerSelection(Slicer slicer, string valueToSelect)
            {
                SlicerCache cache = slicer.SlicerCache;
                foreach (SlicerCacheItem item in cache.SlicerCacheItems)
                {
                    // Deselect all items first
                    item.Selected = false;
                    // Select the desired item
                    if (string.Equals(item.Value, valueToSelect, StringComparison.OrdinalIgnoreCase))
                    {
                        item.Selected = true;
                    }
                }
            }

            // Select "Apple" in the first slicer
            SetSlicerSelection(slicer1, "Apple");

            // Refresh the pivot table so that the slicer filter takes effect
            pivot.RefreshData();
            pivot.CalculateData();

            // Verify selection state of the second slicer
            Console.WriteLine("Selection state of the second slicer after changing the first:");
            foreach (SlicerCacheItem item in slicer2.SlicerCache.SlicerCacheItems)
            {
                Console.WriteLine($"  Item \"{item.Value}\": Selected = {item.Selected}");
            }

            // ------------------------------------------------------------
            // 5. Save the workbook
            // ------------------------------------------------------------
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
