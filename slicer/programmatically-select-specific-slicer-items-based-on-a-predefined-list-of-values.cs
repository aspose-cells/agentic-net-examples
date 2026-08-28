// Title: How to programmatically select specific slicer items in an Aspose.Cells pivot table using C#
// AI Prompts: Write C# code that creates a workbook, adds a pivot table, inserts a slicer for a field, and selects slicer items whose values are contained in a List<string>. | Show how to iterate over SlicerCacheItemCollection in Aspose.Cells and set each item's Selected property based on a predefined collection of values. | Demonstrate refreshing the slicer after updating selections and saving the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# select slicer items from a predefined list | programmatically set slicer selections in a pivot table using Aspose.Cells | C# filter slicer values in Aspose.Cells workbook with specific items | how to update slicer cache items selection in Aspose.Cells C#
// Tags: select slicer items Aspose.Cells C# | slicer cache item manipulation Aspose.Cells | pivot table slicer automation C# | set slicer selected property Aspose.Cells | refresh slicer after selection C#

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerSelectionDemo
{
    // Creates a workbook, adds a pivot table and a slicer, then programmatically selects the slicer items "Apple" and "Banana" based on a predefined List<string>, refreshes the slicer, and saves the file as an .xlsx workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for a pivot table
            cells["A1"].Value = "Fruit";
            cells["B1"].Value = "Sales";
            cells["A2"].Value = "Apple";
            cells["B2"].Value = 120;
            cells["A3"].Value = "Orange";
            cells["B3"].Value = 80;
            cells["A4"].Value = "Banana";
            cells["B4"].Value = 150;
            cells["A5"].Value = "Grape";
            cells["B5"].Value = 60;

            // Add a pivot table based on the data
            int pivotIdx = sheet.PivotTables.Add("A1:B5", "D3", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);      // Fruit column
            pivot.AddFieldToArea(PivotFieldType.Data, 1);     // Sales column
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the "Fruit" field of the pivot table
            int slicerIdx = sheet.Slicers.Add(pivot, "F1", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

            // Predefined list of slicer values that should be selected
            List<string> valuesToSelect = new List<string> { "Apple", "Banana" };

            // Iterate through slicer cache items and set selection based on the list
            SlicerCacheItemCollection cacheItems = slicer.SlicerCache.SlicerCacheItems;
            for (int i = 0; i < cacheItems.Count; i++)
            {
                SlicerCacheItem item = cacheItems[i];
                // Select the item if its value is in the predefined list; otherwise deselect
                item.Selected = valuesToSelect.Contains(item.Value);
            }

            // Refresh the slicer to apply the selection changes
            slicer.Refresh();

            // Save the workbook
            workbook.Save("SlicerSelectionDemo.xlsx");
        }
    }
}
