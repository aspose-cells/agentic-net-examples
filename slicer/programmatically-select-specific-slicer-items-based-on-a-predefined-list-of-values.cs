// Title: C# – Select Specific Slicer Items in an Aspose.Cells Pivot Table from a Predefined List
// Description: Creates a workbook with fruit data, builds a pivot table, adds a slicer linked to the "Fruit" field, and programmatically selects only the slicer items whose values appear in a predefined List<string> (e.g., "Apple" and "Banana"). The slicer cache is updated, refreshed, and the workbook is saved as SlicerSelectionDemo.xlsx.
// Keywords: Aspose.Cells slicer selection C# | programmatic slicer item selection | pivot table slicer cache Aspose | set slicer items from list | refresh slicer Aspose.Cells
// Common Searches: how to select slicer items in Aspose.Cells C# | set slicer selection based on array of values | Aspose.Cells programmatically filter pivot table with slicer | update slicer cache items in .NET | Aspose.Cells select multiple slicer values
// Developer Intent: Automatically select slicer entries that match a given collection of values and deselect every other entry.
// Use Cases: Generate a report that shows only the fruit categories defined by business rules (e.g., Apple and Banana). | Synchronize slicer selections with user‑provided input or an external data source before exporting the workbook. | Reset slicer filters to a default set of items each time a scheduled Excel file is created.
// AI Prompts: Write C# code using Aspose.Cells to select slicer items from a string array and refresh the slicer. | Explain how to loop through SlicerCacheItemCollection and set the Selected property according to a predefined list. | Provide a step‑by‑step guide to programmatically deselect all slicer items except those in a specified list and then save the workbook.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerSelectionDemo
{
    // Creates a workbook with fruit data, builds a pivot table, adds a slicer linked to the "Fruit" field, and programmatically selects only the slicer items whose values appear in a predefined List<string> (e.g., "Apple" and "Banana"). The slicer cache is updated, refreshed, and the workbook is saved as SlicerSelectionDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Predefined list of slicer values to be selected
            List<string> valuesToSelect = new List<string> { "Apple", "Banana" };

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for a pivot table
            cells["A1"].Value = "Fruit";
            cells["B1"].Value = "Quantity";
            cells["A2"].Value = "Apple";
            cells["B2"].Value = 120;
            cells["A3"].Value = "Orange";
            cells["B3"].Value = 80;
            cells["A4"].Value = "Banana";
            cells["B4"].Value = 150;
            cells["A5"].Value = "Grape";
            cells["B5"].Value = 60;

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B5", "D2", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Fruit column
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Quantity column
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the "Fruit" field of the pivot table
            int slicerIdx = sheet.Slicers.Add(pivot, "F2", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

            // Iterate through slicer cache items and set selection based on the predefined list
            SlicerCacheItemCollection cacheItems = slicer.SlicerCache.SlicerCacheItems;
            for (int i = 0; i < cacheItems.Count; i++)
            {
                SlicerCacheItem item = cacheItems[i];
                // Select the item if its value is in the list; otherwise deselect it
                item.Selected = valuesToSelect.Contains(item.Value);
            }

            // Refresh the slicer to apply the selection changes
            slicer.Refresh();

            // Save the workbook
            workbook.Save("SlicerSelectionDemo.xlsx");
        }
    }
}
