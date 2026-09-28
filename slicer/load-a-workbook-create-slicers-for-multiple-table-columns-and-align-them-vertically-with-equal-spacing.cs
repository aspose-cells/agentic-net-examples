// Title: Create and vertically stack slicers for each column of an Excel table using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that adds a slicer for every ListObject column and positions them in a single column with equal vertical gaps. | Show how to calculate and assign TopPixel, LeftPixel, WidthPixel, and HeightPixel for each slicer based on the previous slicer's size in Aspose.Cells. | Explain how to save the workbook after programmatically inserting and aligning multiple slicers in C#.
// Common Searches: aspocells add slicer for each table column c# example | how to align multiple slicers vertically with spacing using Aspose.Cells .NET | set slicer top pixel based on previous slicer Aspose.Cells C# | programmatically create Excel slicers for ListObject columns in C# | vertical slicer layout pixel coordinates Aspose.Cells
// Tags: Aspose.Cells ListObject slicer creation | Aspose.Cells slicer pixel positioning | Aspose.Cells vertical slicer layout | Aspose.Cells multiple slicer alignment | Aspose.Cells C# slicer spacing

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

namespace SlicerAlignmentDemo
{
    // Loads an Excel workbook, creates a table, adds a slicer for each table column, stacks the slicers vertically with equal pixel spacing, and saves the updated file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Create sample data and a table (if the worksheet already has a table, skip this block)
            // ------------------------------------------------------------
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Region");
            cells["C1"].PutValue("Sales");
            cells["A2"].PutValue("A");
            cells["B2"].PutValue("North");
            cells["C2"].PutValue(120);
            cells["A3"].PutValue("B");
            cells["B3"].PutValue("South");
            cells["C3"].PutValue(200);
            cells["A4"].PutValue("A");
            cells["B4"].PutValue("East");
            cells["C4"].PutValue(150);
            cells["A5"].PutValue("B");
            cells["B5"].PutValue("West");
            cells["C5"].PutValue(180);

            // Add a table that covers the data range
            int tableIndex = sheet.ListObjects.Add("A1", "C5", true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // ------------------------------------------------------------
            // Add slicers for each column of the table
            // ------------------------------------------------------------
            SlicerCollection slicers = sheet.Slicers;

            // Configuration for slicer placement
            int startRow = 7;          // Row index (0‑based) where the first slicer will be placed
            int startColumn = 0;       // Column index (0‑based) for all slicers (same column)
            int spacingPixels = 20;    // Space between slicers
            int slicerWidth = 200;     // Width of each slicer in pixels
            int slicerHeight = 150;    // Height of each slicer in pixels
            int leftPixel = 10;        // Horizontal offset for all slicers

            // Loop through each column in the table and create a slicer
            for (int i = 0; i < table.ListColumns.Count; i++)
            {
                // Add slicer using the overload that specifies row and column indices
                int slicerIdx = slicers.Add(table, table.ListColumns[i], startRow + i, startColumn);
                Slicer slicer = slicers[slicerIdx];

                // Set visual properties
                slicer.WidthPixel = slicerWidth;
                slicer.HeightPixel = slicerHeight;
                slicer.LeftPixel = leftPixel;

                // Align vertically: calculate TopPixel based on previous slicer
                if (i == 0)
                {
                    slicer.TopPixel = 10; // Initial top offset
                }
                else
                {
                    Slicer previous = slicers[slicerIdx - 1];
                    slicer.TopPixel = previous.TopPixel + previous.HeightPixel + spacingPixels;
                }

                // Optional: set a caption to identify the slicer
                slicer.Caption = table.ListColumns[i].Name;
            }

            // ------------------------------------------------------------
            // Save the modified workbook
            // ------------------------------------------------------------
            workbook.Save("output.xlsx");
        }
    }
}
