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

            // Sample data: Column A contains categories (the slicer will be linked to this column)
            cells["A1"].Value = "Category";
            cells["A2"].Value = "Fruit";
            cells["A3"].Value = "Fruit";
            cells["A4"].Value = "Vegetable";
            cells["A5"].Value = "Grain";
            cells["A6"].Value = "Fruit";
            cells["A7"].Value = "Grain";

            // Add a pivot table based on the data
            int pivotIdx = sheet.PivotTables.Add("A1:A7", "C3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            // Use the Category column as a row field
            pivot.AddFieldToArea(PivotFieldType.Row, 0);
            // Refresh to populate the pivot
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table for the Category field
            int slicerIdx = sheet.Slicers.Add(pivot, "E3", "Category");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // ------------------------------
            // Dynamically set slicer row height
            // ------------------------------

            // Determine the number of unique items in the source column (Category)
            HashSet<string> uniqueItems = new HashSet<string>();
            // Skip header row (row 0)
            for (int row = 1; row <= sheet.Cells.MaxDataRow; row++)
            {
                object val = cells[row, 0].Value;
                if (val != null)
                {
                    uniqueItems.Add(val.ToString());
                }
            }
            int uniqueCount = uniqueItems.Count;

            // Example logic: base height 15 points + 1 point per unique item
            // Adjust as needed for your UI requirements
            double baseHeight = 15.0;
            slicer.RowHeight = baseHeight + uniqueCount; // height per row in points

            // Optionally, adjust the total slicer height to fit all rows
            slicer.Height = slicer.RowHeight * uniqueCount; // total height in points

            // Save the workbook
            workbook.Save("SlicerDynamicRowHeight.xlsx");
        }
    }
}