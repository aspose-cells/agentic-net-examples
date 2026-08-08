// Title: Sync two Aspose.Cells slicers in C# – automatic selection propagation
// Description: Shows how to build a workbook with a pivot table, add two slicers that filter the same field, select an item in the first slicer, and programmatically copy that selection to the second slicer using a custom SyncSlicers method before refreshing and saving the file.
// Keywords: Aspose.Cells | C# slicer synchronization | pivot table slicer | programmatic slicer selection | Excel slicer linking .NET | duplicate slicer filter | workbook automation | Aspose.Cells API | slicer cache items
// Common Searches: Aspose.Cells sync two slicers | C# copy slicer selection to another slicer | how to link slicers in Aspose.Cells | programmatically update slicer filters | sync slicer selections in .NET workbook
// Developer Intent: Programmatically keep multiple slicers linked to the same pivot field in sync so that a selection change in one slicer automatically updates the others.
// Use Cases: Maintain consistent filtering across dashboard slicers without manual user interaction. | Apply the same filter to slicers placed on different worksheets that reference a shared pivot table. | Refresh slicer states after synchronization to ensure the pivot data reflects the combined filter before exporting.
// AI Prompts: Write a generic C# method that synchronizes any number of Aspose.Cells slicers linked to the same pivot field. | Generate code to handle slicer selection change events and automatically propagate the change to other slicers in real time. | Explain how to preserve slicer synchronization when a workbook is saved, closed, and reopened using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerSyncDemo
{
    // Shows how to build a workbook with a pivot table, add two slicers that filter the same field, select an item in the first slicer, and programmatically copy that selection to the second slicer using a custom SyncSlicers method before refreshing and saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // Populate sample data for the pivot table
            // -------------------------------------------------
            cells["A1"].Value = "Fruit";
            cells["B1"].Value = "Year";
            cells["C1"].Value = "Amount";

            string[] fruits = { "Apple", "Banana", "Cherry", "Apple", "Banana", "Cherry" };
            int[] years = { 2020, 2020, 2020, 2021, 2021, 2021 };
            int[] amounts = { 50, 70, 90, 60, 80, 100 };

            for (int i = 0; i < fruits.Length; i++)
            {
                cells[i + 1, 0].Value = fruits[i];
                cells[i + 1, 1].Value = years[i];
                cells[i + 1, 2].Value = amounts[i];
            }

            // -------------------------------------------------
            // Create a pivot table based on the data range
            // -------------------------------------------------
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIdx = pivots.Add("=Sheet1!A1:C7", "E3", "FruitPivot");
            PivotTable pivot = pivots[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Column, "Year");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // -------------------------------------------------
            // Add two slicers that filter the same pivot field ("Fruit")
            // -------------------------------------------------
            SlicerCollection slicers = sheet.Slicers;

            // First slicer placed at G3
            int slicerIdx1 = slicers.Add(pivot, "G3", "Fruit");
            Slicer slicer1 = slicers[slicerIdx1];
            slicer1.Caption = "Fruit Slicer A";

            // Second slicer placed at G15
            int slicerIdx2 = slicers.Add(pivot, "G15", "Fruit");
            Slicer slicer2 = slicers[slicerIdx2];
            slicer2.Caption = "Fruit Slicer B";

            // -------------------------------------------------
            // Example: select "Banana" in the first slicer
            // -------------------------------------------------
            SelectSlicerItem(slicer1, "Banana");

            // -------------------------------------------------
            // Synchronize the second slicer with the first slicer
            // -------------------------------------------------
            SyncSlicers(slicer1, slicer2);

            // Refresh both slicers to reflect the changes in the pivot table
            slicer1.Refresh();
            slicer2.Refresh();

            // Save the workbook
            workbook.Save("SlicerSyncDemo.xlsx");
        }

        static void SelectSlicerItem(Slicer slicer, string itemValue)
        {
            foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
            {
                // Select the matching item, deselect the rest
                item.Selected = string.Equals(item.Value, itemValue, StringComparison.OrdinalIgnoreCase);
            }
        }

        /// <summary>
        /// Copies the selection state from a source slicer to a target slicer.
        /// Both slicers must be based on the same field (same cache items).
        /// </summary>
        static void SyncSlicers(Slicer source, Slicer target)
        {
            // Build a lookup of source selections by item value
            var selectedValues = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (SlicerCacheItem srcItem in source.SlicerCache.SlicerCacheItems)
            {
                if (srcItem.Selected)
                    selectedValues.Add(srcItem.Value);
            }

            // Apply the same selection to the target slicer
            foreach (SlicerCacheItem tgtItem in target.SlicerCache.SlicerCacheItems)
            {
                tgtItem.Selected = selectedValues.Contains(tgtItem.Value);
            }
        }
    }
}
