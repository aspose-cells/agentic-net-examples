// Title: Aspose.Cells .NET – Apply a Label Filter and a Top 10 Filter to the Same Pivot Field
// Description: Demonstrates how to enable multiple filters on a pivot field, add a CaptionBeginsWith label filter for items starting with "A" and a Top 10 count filter for the top two categories, refresh the pivot cache, recalculate, and save the workbook using Aspose.Cells for C#.
// Keywords: Aspose.Cells C# pivot filter | label filter pivot table .NET | Top10 filter Aspose.Cells | AllowMultipleFiltersPerField | combined pivot filters | C# Excel automation | global developers | US .NET community | Europe Excel libraries
// Common Searches: how to add both label and top10 filters to one pivot field in Aspose.Cells | enable multiple filters per pivot field C# | Aspose.Cells CaptionBeginsWith filter example | Top10 count filter on pivot table programmatically | refresh pivot cache after adding filters Aspose.Cells
// Developer Intent: Programmatically apply a caption‑begins‑with label filter and a Top 10 count filter to the same pivot field in a .NET workbook.
// Use Cases: Show only product categories that start with a specific letter while limiting the view to the top N categories by transaction count. | Create compliance‑driven reports that require both a textual prefix filter and a ranking filter on the same dimension. | Automate Excel dashboards where multiple pivot filters must be applied before publishing the file.
// AI Prompts: Write C# code with Aspose.Cells to add a CaptionBeginsWith label filter and a Top10 count filter to the same pivot field, ensuring AllowMultipleFiltersPerField is enabled. | Explain the role of AllowMultipleFiltersPerField in Aspose.Cells and give an example of combining a label filter with a Top10 filter. | Modify the sample to filter categories starting with "B" and display the top 5 items by sum instead of count.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to enable multiple filters on a pivot field, add a CaptionBeginsWith label filter for items starting with "A" and a Top 10 count filter for the top two categories, refresh the pivot cache, recalculate, and save the workbook using Aspose.Cells for C#.
    public class PivotFieldCombinedFiltersDemo
    {
        public static void Main(string[] args)
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

            // Populate sample data for the pivot table
            // Column A: Category, Column B: Value
            cells["A1"].Value = "Category";
            cells["A2"].Value = "Apple";
            cells["A3"].Value = "Banana";
            cells["A4"].Value = "Avocado";
            cells["A5"].Value = "Cherry";
            cells["A6"].Value = "Apricot";

            cells["B1"].Value = "Value";
            cells["B2"].Value = 120;
            cells["B3"].Value = 80;
            cells["B4"].Value = 150;
            cells["B5"].Value = 60;
            cells["B6"].Value = 200;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add the Category field as a row field and Value as a data field
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Enable multiple filters on the same pivot field
            pivot.AllowMultipleFiltersPerField = true;

            // Get the collection of pivot filters
            PivotFilterCollection filters = pivot.PivotFilters;

            // -------------------------------------------------
            // 1. Add a label filter: show categories that begin with "A"
            // -------------------------------------------------
            PivotFilter labelFilter = filters.AddLabelFilter(
                baseFieldIndex: 0,                     // Index of the Category field in source data
                type: PivotFilterType.CaptionBeginsWith,
                label1: "A",
                label2: null);

            // -------------------------------------------------
            // 2. Add a Top10 filter: show top 2 categories by count
            // -------------------------------------------------
            PivotFilter top10Filter = filters.AddTop10Filter(
                baseFieldIndex: 0,                     // Category field
                valueFieldIndex: 1,                    // Value field (data field index)
                type: PivotFilterType.Count,           // Filter based on count
                isTop: true,                           // Top items
                itemCount: 2);                         // Show top 2 items

            // Refresh pivot cache and calculate the pivot table to apply filters
            pivot.RefreshData();      // Correct method to refresh the cache
            pivot.CalculateData();    // Recalculate after applying filters

            // Save the workbook with the applied filters
            workbook.Save("PivotFieldCombinedFiltersDemo.xlsx");
        }
    }
}
