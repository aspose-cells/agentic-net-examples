using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

namespace SlicerAlignmentDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one table; if not, create a simple sample table
            if (sheet.ListObjects.Count == 0)
            {
                // Sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("A");
                sheet.Cells["B4"].PutValue(30);

                // Add the table covering the sample range
                int tableIdx = sheet.ListObjects.Add("A1", "B4", true);
                // Optional: apply a style
                sheet.ListObjects[tableIdx].TableStyleType = TableStyleType.TableStyleMedium2;
            }

            // Reference the first table on the worksheet
            ListObject table = sheet.ListObjects[0];

            // Get the slicer collection for the worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Configuration for vertical alignment
            int startRow = 5;               // Row index for the first slicer (0‑based)
            int startColumn = 5;            // Column index for all slicers (0‑based)
            int slicerHeight = 150;         // Height of each slicer in pixels
            int slicerWidth = 200;          // Width of each slicer in pixels
            int verticalSpacing = 10;       // Space between slicers in pixels

            // Create a slicer for each column of the table and align them vertically
            for (int i = 0; i < table.ListColumns.Count; i++)
            {
                // Add slicer using the overload that accepts ListObject, ListColumn, row, column
                int slicerIdx = slicers.Add(table, table.ListColumns[i], startRow, startColumn);
                Slicer slicer = slicers[slicerIdx];

                // Set size
                slicer.HeightPixel = slicerHeight;
                slicer.WidthPixel = slicerWidth;

                // Position slicer vertically with equal spacing
                slicer.TopPixel = i * (slicerHeight + verticalSpacing);
                slicer.LeftPixel = 50; // Fixed left offset for all slicers

                // Optional: give each slicer a meaningful name
                slicer.Name = $"Slicer_{table.ListColumns[i].Name}";
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}