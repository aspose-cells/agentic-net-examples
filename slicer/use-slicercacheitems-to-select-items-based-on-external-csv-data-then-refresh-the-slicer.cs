// Title: Aspose.Cells C# – Select Slicer Items from CSV and Refresh Pivot
// Description: Creates a workbook with fruit sales, adds a pivot table and slicer, reads fruit names from a CSV, selects matching slicer items, refreshes the slicer (updating the pivot), and saves the file.
// Keywords: Aspose.Cells | C# | .NET | slicer | SlicerCacheItem | CSV | pivot table | programmatic selection | refresh slicer | Excel automation
// Common Searches: Aspose.Cells select slicer items from CSV | C# set slicer selections programmatically | Refresh pivot after slicer update Aspose.Cells | Load filter list CSV into slicer .NET | How to use SlicerCacheItem.Selected Aspose.Cells
// Developer Intent: Programmatically set slicer selections based on CSV data and refresh the linked pivot table using Aspose.Cells for .NET.
// Use Cases: Apply a dynamic filter to a pivot table by loading category names from a CSV and marking the corresponding slicer items as selected. | Synchronize slicer selections across multiple reports by exporting chosen items to CSV and re‑applying them in another workbook. | Automate report generation where filter criteria are supplied externally in a CSV file, ensuring the pivot reflects those selections.
// AI Prompts: Generate C# code with Aspose.Cells that reads a CSV of product names, selects matching slicer items, refreshes the slicer, and saves the workbook. | Show how to deselect all slicer items except those listed in an external text file, then refresh the linked pivot table using Aspose.Cells. | Explain case‑insensitive matching when updating SlicerCacheItem.Selected from CSV values in Aspose.Cells for .NET.

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerCsvDemo
{
    // Creates a workbook with fruit sales, adds a pivot table and slicer, reads fruit names from a CSV, selects matching slicer items, refreshes the slicer (updating the pivot), and saves the file.
    class Program
    {
        static void Main()
        {
            // ---------- Create a workbook and sample data ----------
            Workbook workbook = new Workbook(); // create
            Worksheet dataSheet = workbook.Worksheets[0];
            Cells cells = dataSheet.Cells;

            // Sample data for pivot table (Fruit, Sales)
            cells["A1"].PutValue("Fruit");
            cells["B1"].PutValue("Sales");
            string[] fruits = { "Apple", "Banana", "Orange", "Grape", "Kiwi" };
            int[] sales = { 120, 80, 150, 60, 90 };
            for (int i = 0; i < fruits.Length; i++)
            {
                cells[i + 1, 0].PutValue(fruits[i]);   // Column A
                cells[i + 1, 1].PutValue(sales[i]);   // Column B
            }

            // ---------- Create a pivot table ----------
            int pivotIndex = dataSheet.PivotTables.Add("A1:B6", "D3", "FruitPivot");
            PivotTable pivot = dataSheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Fruit column
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Sales column
            pivot.RefreshData();
            pivot.CalculateData();

            // ---------- Add a slicer linked to the pivot table ----------
            int slicerIndex = dataSheet.Slicers.Add(pivot, "F3", "Fruit");
            Slicer slicer = dataSheet.Slicers[slicerIndex];
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

            // ---------- Load external CSV containing items to select ----------
            // CSV format: one column with fruit names to be selected, e.g.:
            // Apple
            // Orange
            string csvPath = "filter.csv";
            HashSet<string> itemsToSelect = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (File.Exists(csvPath))
            {
                foreach (string line in File.ReadAllLines(csvPath))
                {
                    string trimmed = line.Trim();
                    if (!string.IsNullOrEmpty(trimmed))
                        itemsToSelect.Add(trimmed);
                }
            }

            // ---------- Set slicer cache items based on CSV ----------
            SlicerCacheItemCollection cacheItems = slicer.SlicerCache.SlicerCacheItems;
            for (int i = 0; i < cacheItems.Count; i++)
            {
                SlicerCacheItem item = cacheItems[i];
                // Select the item if its value exists in the CSV list; otherwise deselect
                item.Selected = itemsToSelect.Contains(item.Value);
            }

            // ---------- Refresh the slicer (also refreshes the pivot table) ----------
            slicer.Refresh();

            // ---------- Save the workbook ----------
            workbook.Save("SlicerCsvDemo.xlsx"); // save
        }
    }
}
