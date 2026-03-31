using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

namespace SlicerDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            string inputPath = "InputWorkbook.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (or any worksheet you need)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Ensure there is a data source for the slicer.
            // Here we create a simple table if it does not already exist.
            // ------------------------------------------------------------
            // Populate sample data if the sheet is empty
            if (cells.MaxDataRow < 0 || cells.MaxDataColumn < 0)
            {
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue("A");
                cells["B2"].PutValue(10);
                cells["A3"].PutValue("B");
                cells["B3"].PutValue(20);
                cells["A4"].PutValue("C");
                cells["B4"].PutValue(30);
            }

            // Add a ListObject (table) covering the data range
            int startRow = 0;
            int startColumn = 0;
            int endRow = cells.MaxDataRow;
            int endColumn = cells.MaxDataColumn;
            int tableIndex = sheet.ListObjects.Add(startRow, startColumn, endRow, endColumn, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // ------------------------------------------------------------
            // Add a slicer linked to the first column of the table.
            // The slicer will be placed with its upper‑left corner at cell E1.
            // ------------------------------------------------------------
            SlicerCollection slicers = sheet.Slicers;
            int slicerIdx = slicers.Add(table, 0, "E1"); // 0 = first column in the table
            Slicer slicer = slicers[slicerIdx];

            // Optional: set some slicer properties
            slicer.Caption = "Category Slicer";
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

            // ------------------------------------------------------------
            // Remove the slicer that was just added.
            // Demonstrate both removal methods.
            // ------------------------------------------------------------
            // Method 1: Remove by index
            slicers.RemoveAt(slicerIdx);

            // If you prefer to remove by object reference, you could use:
            // slicers.Remove(slicer);

            // ------------------------------------------------------------
            // Save the modified workbook.
            // ------------------------------------------------------------
            string outputPath = "OutputWorkbook.xlsx";
            workbook.Save(outputPath);
        }
    }
}