// Title: Enable multi‑select on a slicer linked to a pivot table with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells enable slicer multi select | How to programmatically select multiple items in an Aspose.Cells slicer | Aspose.Cells slicer multiple selection example | Set slicer to allow multiple categories in pivot table using Aspose.Cells | Aspose.Cells C# multi‑select slicer for pivot table
// Tags: Aspose.Cells slicer multi‑select implementation | pivot table slicer cache manipulation C# | C# programmatic slicer item selection | Excel workbook slicer multi‑choice Aspose | Aspose.Cells create slicer for pivot table

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

// This C# example demonstrates creating a workbook, converting a data range to a table, adding a pivot table, inserting a slicer linked to the pivot's Category field, and programmatically selecting multiple slicer items by setting the Selected property on SlicerCacheItem objects, resulting in a multi‑select slicer saved as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // -----------------------------------------------------------------
            // 1. Create a workbook and fill it with sample data
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Year");
            cells["C1"].PutValue("Sales");

            // Data rows
            string[] categories = { "Apple", "Banana", "Cherry", "Date", "Elderberry" };
            int[] years = { 2020, 2021 };
            int row = 1;
            Random rnd = new Random();
            foreach (var cat in categories)
            {
                foreach (var yr in years)
                {
                    cells[row, 0].PutValue(cat);
                    cells[row, 1].PutValue(yr);
                    cells[row, 2].PutValue(rnd.Next(50, 200));
                    row++;
                }
            }

            // -----------------------------------------------------------------
            // 2. Create a table from the data range (required for slicer)
            // -----------------------------------------------------------------
            int lastRow = row - 1; // zero‑based index of the last data row
            int tableIndex = sheet.ListObjects.Add(0, 0, lastRow, 2, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // -----------------------------------------------------------------
            // 3. Add a pivot table based on the table
            // -----------------------------------------------------------------
            // Build the source range string (e.g., "A1:C11")
            string sourceRange = $"=Sheet1!A1:C{lastRow + 1}";
            int pivotIndex = sheet.PivotTables.Add(sourceRange, "E3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Column, "Year");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;

            // Refresh and calculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // -----------------------------------------------------------------
            // 4. Add a slicer linked to the pivot table (field: Category)
            // -----------------------------------------------------------------
            int slicerIndex = sheet.Slicers.Add(pivot, "Category", "F2");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // -----------------------------------------------------------------
            // 5. (Optional) Pre‑select a few items programmatically
            // -----------------------------------------------------------------
            SlicerCache cache = slicer.SlicerCache;
            SlicerCacheItemCollection items = cache.SlicerCacheItems;

            // Deselect all first (ensures a clean state)
            foreach (SlicerCacheItem item in items)
                item.Selected = false;

            // Select the first two items (if they exist)
            if (items.Count >= 2)
            {
                items[0].Selected = true; // e.g., "Apple"
                items[1].Selected = true; // e.g., "Banana"
            }

            // -----------------------------------------------------------------
            // 6. Save the workbook
            // -----------------------------------------------------------------
            string outPath = "MultiSelectSlicerDemo.xlsx";
            workbook.Save(outPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
