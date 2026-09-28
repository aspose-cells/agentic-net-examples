// Title: Reset an Excel slicer to its default state by selecting all items with Aspose.Cells for .NET (C#)
// AI Prompts: Provide C# code that iterates through a slicer's SlicerCacheItems, sets each item's Selected property to true, and refreshes the slicer using Aspose.Cells. | Show how to programmatically clear a slicer filter in an Excel workbook by marking all slicer items as selected with Aspose.Cells for .NET. | Generate a method that resets a pivot table slicer to its initial state in a .NET workbook and saves the file.
// Common Searches: Aspose.Cells C# clear slicer selections programmatically | How to remove slicer filter from an Excel workbook using Aspose.Cells | Select every item in an Excel slicer with C# and Aspose.Cells | Refresh a slicer after updating its items in a .NET workbook | Reset pivot table slicer to show all data via code in C#
// Tags: Aspose.Cells C# slicer selection reset | iterate slicer cache items .NET | apply slicer refresh Aspose.Cells | pivot table slicer default configuration | remove slicer filter programmatically Excel

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerReset
{
    // The example creates a workbook, adds sample data, builds a pivot table, attaches a slicer to the 'Category' field, initially selects only the first slicer item, then resets the slicer by marking every SlicerCacheItem as selected, refreshes the slicer, and saves the workbook as SlicerResetDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Category";
            cells["A2"].Value = "A";
            cells["A3"].Value = "B";
            cells["A4"].Value = "C";
            cells["B1"].Value = "Value";
            cells["B2"].Value = 10;
            cells["B3"].Value = 20;
            cells["B4"].Value = 30;

            // Add a pivot table based on the data
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table
            SlicerCollection slicers = sheet.Slicers;
            int slicerIndex = slicers.Add(pivot, "F1", "Category");
            Slicer slicer = slicers[slicerIndex];

            // Example: select only the first item (simulating a filter)
            for (int i = 0; i < slicer.SlicerCache.SlicerCacheItems.Count; i++)
            {
                slicer.SlicerCache.SlicerCacheItems[i].Selected = i == 0;
            }
            slicer.Refresh();

            // ----- Reset the slicer to its default state -----
            // Set all items as selected so that no filter is applied
            foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
            {
                item.Selected = true;
            }
            // Refresh the slicer to apply the changes
            slicer.Refresh();

            // Save the workbook
            workbook.Save("SlicerResetDemo.xlsx");
        }
    }
}
