using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDynamicRowHeight
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data in column A (the linked column for the slicer)
            cells["A1"].Value = "Category";
            cells["A2"].Value = "Fruit";
            cells["A3"].Value = "Fruit";
            cells["A4"].Value = "Vegetable";
            cells["A5"].Value = "Grain";
            cells["A6"].Value = "Fruit";
            cells["A7"].Value = "Grain";

            // Create a pivot table based on the data
            int pivotIdx = sheet.PivotTables.Add("A1:A7", "C3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0); // use column A as row field

            // Add a slicer linked to the pivot table for the "Category" field
            int slicerIdx = sheet.Slicers.Add(pivot, "E3", "Category");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // ----- Dynamic RowHeight calculation -----
            // Determine the number of unique items in the linked column (A)
            HashSet<string> uniqueItems = new HashSet<string>();
            for (int row = 1; row <= cells.MaxDataRow; row++) // start from row 2 (index 1) to skip header
            {
                object val = cells[row, 0].Value;
                if (val != null)
                    uniqueItems.Add(val.ToString());
            }
            int uniqueCount = uniqueItems.Count;

            // Example logic: smaller row height for many items, larger for few items
            // Adjust the per‑row height (in points)
            if (uniqueCount > 10)
                slicer.RowHeight = 12;   // compact rows
            else if (uniqueCount > 5)
                slicer.RowHeight = 16;
            else
                slicer.RowHeight = 20;   // spacious rows

            // Optionally adjust the overall slicer height so all items are visible
            slicer.Height = slicer.RowHeight * uniqueCount;

            // Save the workbook
            workbook.Save("SlicerDynamicRowHeight.xlsx");
        }
    }
}