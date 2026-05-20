using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;
using Aspose.Cells.Drawing;

namespace SlicerPlacementExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the table (two columns: Category and Value)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("C");
            sheet.Cells["B5"].PutValue(40);

            // Add a ListObject (table) that covers the data range A1:B5
            int tableIndex = sheet.ListObjects.Add(0, 0, 4, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.TableStyleType = TableStyleType.TableStyleMedium2;

            // Add a slicer linked to the first column ("Category") of the table.
            // The slicer will be placed starting at row 7, column 5 (cell E7).
            SlicerCollection slicers = sheet.Slicers;
            int slicerIndex = slicers.Add(table, table.ListColumns[0], 6, 4); // zero‑based indices
            Slicer slicer = slicers[slicerIndex];

            // Set the slicer title (optional)
            slicer.Title = "Category Filter";

            // Set the placement so the slicer moves and resizes with the cells.
            // This uses the obsolete Placement property as required.
            slicer.Placement = PlacementType.MoveAndSize;

            // Position the slicer at the top‑right corner of the worksheet.
            // TopPixel = 0 places it at the top; LeftPixel is set to a large value
            // to push it towards the right edge (adjust as needed for your sheet size).
            slicer.TopPixel = 0;
            slicer.LeftPixel = 800; // approximate right‑most position

            // Optionally adjust size
            slicer.WidthPixel = 150;
            slicer.HeightPixel = 200;

            // Save the workbook
            workbook.Save("SlicerTopRightPlacement.xlsx");
        }
    }
}