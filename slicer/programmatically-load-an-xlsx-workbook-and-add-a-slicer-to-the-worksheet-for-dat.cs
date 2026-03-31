using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

namespace SlicerExample
{
    class Program
    {
        static void Main()
        {
            // Paths for the source and result workbooks
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (you can change the index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Ensure there is a table (ListObject) to which the slicer can bind.
            // Here we create a table covering the range A1:B5.
            // Adjust the range according to your actual data layout.
            // ------------------------------------------------------------
            int firstRow = 0;   // A
            int firstColumn = 0; // 1
            int lastRow = 4;    // row 5 (zero‑based)
            int lastColumn = 1; // column B
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, lastRow, lastColumn, true);
            ListObject table = sheet.ListObjects[tableIndex];

            // ------------------------------------------------------------
            // Add a slicer for the first column of the table.
            // The slicer will be placed with its upper‑left corner at cell E2.
            // ------------------------------------------------------------
            int slicerIndex = sheet.Slicers.Add(table, 0, "E2");
            Slicer slicer = sheet.Slicers[slicerIndex];

            // Optional: set a caption for the slicer
            slicer.Caption = "Category Filter";

            // Save the modified workbook
            workbook.Save(outputPath);
        }
    }
}