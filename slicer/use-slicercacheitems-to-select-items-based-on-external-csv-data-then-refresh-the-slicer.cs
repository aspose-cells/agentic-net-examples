// Title: Programmatically select Excel slicer items from a CSV file and refresh the linked pivot table using Aspose.Cells for .NET (C#)
// AI Prompts: Load a list of values from a CSV file and set the matching SlicerCacheItem.Selected flags in an Aspose.Cells workbook. | Apply external CSV data to an Excel slicer, then call slicer.Refresh to update the connected pivot table. | Create a pivot table with a slicer, read filter criteria from a CSV, programmatically select the corresponding slicer items, and save the workbook.
// Common Searches: Aspose.Cells C# select slicer items using values from a CSV file | How to refresh an Excel slicer after changing SlicerCacheItem.Selected in .NET | Programmatic Excel slicer filtering with external CSV data using Aspose.Cells | C# example linking a slicer to a pivot table and applying CSV‑based filters
// Tags: Aspose.Cells slicer cache item selection from CSV | C# programmatic slicer filtering in Excel | refresh slicer linked pivot table Aspose.Cells | load external filter values into Excel slicer | pivot table slicer automation with Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerCsvDemo
{
    // // Demonstrates creating a workbook, adding a pivot table and slicer, loading filter values from a CSV file, selecting matching slicer cache items, refreshing the slicer (and linked pivot), and saving the workbook.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].Value = "Fruit";
            dataSheet.Cells["A2"].Value = "Apple";
            dataSheet.Cells["A3"].Value = "Orange";
            dataSheet.Cells["A4"].Value = "Banana";
            dataSheet.Cells["A5"].Value = "Apple";
            dataSheet.Cells["A6"].Value = "Banana";

            dataSheet.Cells["B1"].Value = "Sales";
            dataSheet.Cells["B2"].Value = 120;
            dataSheet.Cells["B3"].Value = 150;
            dataSheet.Cells["B4"].Value = 200;
            dataSheet.Cells["B5"].Value = 130;
            dataSheet.Cells["B6"].Value = 210;

            // ---------- Create a pivot table ----------
            int pivotIdx = dataSheet.PivotTables.Add("A1:B6", "D2", "FruitPivot");
            PivotTable pivot = dataSheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.RefreshData();
            pivot.CalculateData();

            // ---------- Add a slicer linked to the pivot ----------
            int slicerIdx = dataSheet.Slicers.Add(pivot, "F2", "Fruit");
            Slicer slicer = dataSheet.Slicers[slicerIdx];
            slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

            // ---------- Load external CSV containing values to be selected ----------
            // CSV format: each line contains a fruit name to select, e.g.
            // Apple
            // Banana
            string csvPath = "filter.csv";
            HashSet<string> valuesToSelect = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (File.Exists(csvPath))
            {
                foreach (string line in File.ReadAllLines(csvPath))
                {
                    string trimmed = line.Trim();
                    if (!string.IsNullOrEmpty(trimmed))
                        valuesToSelect.Add(trimmed);
                }
            }

            // ---------- Set slicer cache items selection based on CSV ----------
            SlicerCacheItemCollection cacheItems = slicer.SlicerCache.SlicerCacheItems;
            for (int i = 0; i < cacheItems.Count; i++)
            {
                SlicerCacheItem item = cacheItems[i];
                // Select the item if its value exists in the CSV list
                item.Selected = valuesToSelect.Contains(item.Value);
            }

            // ---------- Refresh the slicer (also refreshes the linked pivot table) ----------
            slicer.Refresh();

            // ---------- Save the workbook ----------
            workbook.Save("SlicerCsvDemo.xlsx");
        }
    }
}
