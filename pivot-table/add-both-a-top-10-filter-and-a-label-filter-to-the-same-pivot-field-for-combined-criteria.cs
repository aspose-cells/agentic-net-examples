// Title: Aspose.Cells .NET – Apply Top 10 and Label Filters to the Same Pivot Field
// Description: Demonstrates how to create a workbook, build a pivot table, enable multiple filters per field, and add both a Top 10 filter (showing the top 2 categories by item count) and a label filter (excluding blank categories) to the same row field. The example refreshes, calculates, and saves the workbook using C# and Aspose.Cells.
// Keywords: Aspose.Cells | C# | PivotTable | Top10 filter | Label filter | multiple filters per field | AllowMultipleFiltersPerField | exclude blank categories | combined pivot filters | Excel automation | Aspose.Cells .NET tutorial | US developers | UK developers | India developers
// Common Searches: Aspose.Cells add Top10 filter and label filter to same pivot field | AllowMultipleFiltersPerField example C# | exclude blank rows in Aspose.Cells pivot table | combined Top N and label filters Aspose.Cells | pivot table multiple filters Aspose.Cells .NET
// Developer Intent: Add a Top 10 filter and a label filter to the same pivot field in a workbook using Aspose.Cells for .NET.
// Use Cases: Show only the two categories with the highest sales count while removing blank category rows. | Create a pivot report that simultaneously applies a Top‑N filter and a not‑equal label filter on a single field. | Generate an Excel file where a row field is limited to top items and excludes specific labels in one step.
// AI Prompts: Generate C# code with Aspose.Cells that adds a Top10 filter (by count) and a CaptionNotEqual label filter to the same row field of a pivot table. | Explain the purpose of AllowMultipleFiltersPerField and the required steps to refresh and calculate a pivot table after adding combined filters. | Provide a step‑by‑step guide to filter out blank categories and keep the top 3 products in a pivot table using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, build a pivot table, enable multiple filters per field, and add both a Top 10 filter (showing the top 2 categories by item count) and a label filter (excluding blank categories) to the same row field. The example refreshes, calculates, and saves the workbook using C# and Aspose.Cells.
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
                cells["A1"].Value = "Category";
                cells["A2"].Value = "A";
                cells["A3"].Value = "B";
                cells["A4"].Value = "C";
                cells["A5"].Value = "A";
                cells["A6"].Value = "B";
                cells["A7"].Value = "C";
                cells["A8"].Value = "";          // blank category
                cells["B1"].Value = "Sales";
                cells["B2"].Value = 120;
                cells["B3"].Value = 80;
                cells["B4"].Value = 150;
                cells["B5"].Value = 200;
                cells["B6"].Value = 90;
                cells["B7"].Value = 110;
                cells["B8"].Value = 50;

                // Create a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B8", "D3", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIndex];

                // Add the "Category" field as a row field and "Sales" as a data field
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Enable multiple filters on the same pivot field
                pivot.AllowMultipleFiltersPerField = true;

                // Get the collection of pivot filters
                PivotFilterCollection filters = pivot.PivotFilters;

                // -------------------------------------------------
                // 1. Add a Top10 filter: show top 2 categories by count
                // -------------------------------------------------
                // baseFieldIndex = 0 (Category field)
                // valueFieldIndex = 1 (Sales data field)
                // Filter type = Count (filter by item count)
                // isTop = true (top items), itemCount = 2
                filters.AddTop10Filter(
                    baseFieldIndex: 0,
                    valueFieldIndex: 1,
                    type: PivotFilterType.Count,
                    isTop: true,
                    itemCount: 2);

                // -------------------------------------------------
                // 2. Add a label filter: exclude blank categories
                // -------------------------------------------------
                // type = CaptionNotEqual, label1 = "(blank)", label2 = "" (empty string)
                filters.AddLabelFilter(
                    baseFieldIndex: 0,
                    type: PivotFilterType.CaptionNotEqual,
                    label1: "(blank)",
                    label2: "");

                // Refresh and calculate the pivot table to apply filters
                pivot.RefreshData();
                pivot.CalculateData();

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
