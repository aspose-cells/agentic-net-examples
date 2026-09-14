// Title: How to apply both a CaptionBeginsWith label filter and a Top 10 count filter to the same pivot field using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a pivot table, enables multiple filters per field, adds a label filter for rows whose text starts with 'A', then adds a Top 10 count filter to keep the two highest‑valued categories, refreshes the pivot cache and saves the workbook. | Write a C# example using Aspose.Cells to apply both a starts‑with‑letter label filter and a top‑N count filter on the same row field of a pivot table, including setting AllowMultipleFiltersPerField and updating the pivot data.
// Common Searches: Aspose.Cells C# combine label filter and top 10 filter on same pivot field | Enable multiple filters per pivot field in Aspose.Cells .NET | CaptionBeginsWith filter example Aspose.Cells pivot table | Top10 count filter on pivot table using Aspose.Cells C# | Refresh pivot cache after applying filters Aspose.Cells
// Tags: combined label and top10 filters Aspose.Cells | label filter starts with A Aspose.Cells | top10 count filter Aspose.Cells | AllowMultipleFiltersPerField Aspose.Cells | refresh pivot cache Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data, builds a pivot table, enables multiple filters per field, applies a label filter that shows categories beginning with "A" and a Top 10 count filter that displays the top two categories, refreshes and recalculates the pivot, and saves the result as PivotFieldCombinedFiltersDemo.xlsx.
    public class PivotFieldCombinedFiltersDemo
    {
        public static void Run()
        {
            try
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
                cells["A4"].Value = "Apricot";
                cells["A5"].Value = "Blueberry";
                cells["A6"].Value = "Avocado";

                cells["B1"].Value = "Value";
                cells["B2"].Value = 120;
                cells["B3"].Value = 80;
                cells["B4"].Value = 150;
                cells["B5"].Value = 60;
                cells["B6"].Value = 200;

                // Create a pivot table based on the data range
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
                filters.AddLabelFilter(
                    baseFieldIndex: 0,                     // Index of the Category field in the source
                    type: PivotFilterType.CaptionBeginsWith,
                    label1: "A",
                    label2: null);

                // -------------------------------------------------
                // 2. Add a Top10 filter: show top 2 categories by count
                // -------------------------------------------------
                filters.AddTop10Filter(
                    baseFieldIndex: 0,                     // Category field
                    valueFieldIndex: 1,                    // Value field (data field)
                    type: PivotFilterType.Count,
                    isTop: true,                           // Top items
                    itemCount: 2);                         // Show top 2

                // Refresh the pivot cache and recalculate the pivot table to apply filters
                pivot.RefreshData();      // Refresh pivot cache
                pivot.CalculateData();    // Recalculate pivot table

                // Save the workbook with the applied filters
                workbook.Save("PivotFieldCombinedFiltersDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            PivotFieldCombinedFiltersDemo.Run();
        }
    }
}
