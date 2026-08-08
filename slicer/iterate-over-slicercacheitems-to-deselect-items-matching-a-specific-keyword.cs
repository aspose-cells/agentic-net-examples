// Title: C# – Deselect Pivot Table Slicer Items by Keyword with Aspose.Cells
// Description: Shows how to build a workbook, add a pivot table and a linked slicer, then walk through the slicer's SlicerCacheItems to unselect any item whose value matches a supplied keyword (case‑insensitive). The slicer is refreshed and the workbook saved.
// Keywords: Aspose.Cells | C# slicer | pivot table slicer | SlicerCacheItem | deselect slicer item | keyword filter | case insensitive selection | programmatic slicer | Excel automation | Aspose.Cells API
// Common Searches: Aspose.Cells deselect slicer item C# | How to unselect slicer values in Aspose.Cells | Iterate SlicerCacheItems Aspose.Cells | Remove specific slicer selection programmatically | Case‑insensitive slicer filter Aspose.Cells
// Developer Intent: Remove slicer selections that match a specific text value in a pivot‑table slicer using Aspose.Cells for .NET.
// Use Cases: Generate reports that automatically exclude a particular category (e.g., a fruit) without manual interaction. | Automate data cleansing by programmatically turning off unwanted slicer options before publishing. | Create dynamic dashboards where slicer selections are preset based on business rules. | Prepare workbooks for distribution with sensitive items hidden via slicer deselection.
// AI Prompts: Write C# code with Aspose.Cells to loop through a slicer's SlicerCacheItems and set Selected = false for items equal to a supplied keyword, ignoring case. | Show how to refresh a slicer after changing its cache items in Aspose.Cells. | Explain the steps to link a slicer to a pivot table field and programmatically filter its items by string value in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsExamples
{
    // Shows how to build a workbook, add a pivot table and a linked slicer, then walk through the slicer's SlicerCacheItems to unselect any item whose value matches a supplied keyword (case‑insensitive). The slicer is refreshed and the workbook saved.
    public class DeselectSlicerItemsByKeyword
    {
        // Entry point for the application
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for a pivot table
            cells["A1"].Value = "Fruit";
            cells["A2"].Value = "Apple";
            cells["A3"].Value = "Orange";
            cells["A4"].Value = "Banana";
            cells["A5"].Value = "Apple";
            cells["B1"].Value = "Quantity";
            cells["B2"].Value = 10;
            cells["B3"].Value = 20;
            cells["B4"].Value = 15;
            cells["B5"].Value = 12;

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B5", "D1", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the "Fruit" field of the pivot table
            SlicerCollection slicers = sheet.Slicers;
            int slicerIdx = slicers.Add(pivot, "Fruit", "FruitSlicer");
            Slicer slicer = slicers[slicerIdx];
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

            // Define the keyword for which items should be deselected
            string keyword = "Apple";

            // Deselect slicer items that match the keyword (case‑insensitive)
            SlicerCacheItemCollection cacheItems = slicer.SlicerCache.SlicerCacheItems;
            foreach (SlicerCacheItem item in cacheItems)
            {
                if (string.Equals(item.Value, keyword, StringComparison.OrdinalIgnoreCase))
                {
                    item.Selected = false;
                }
            }

            // Refresh the slicer to apply the selection changes
            slicer.Refresh();

            // Save the workbook
            workbook.Save("DeselectSlicerItemsByKeyword.xlsx");
        }
    }
}
