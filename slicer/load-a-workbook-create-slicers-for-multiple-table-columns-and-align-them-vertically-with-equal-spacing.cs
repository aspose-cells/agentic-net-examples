using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet contains at least one table
        if (worksheet.ListObjects.Count == 0)
        {
            Console.WriteLine("No tables found in the worksheet.");
            return;
        }

        // Use the first table in the worksheet as the data source for slicers
        ListObject table = worksheet.ListObjects[0];

        // Define which columns of the table will have slicers (zero‑based indices)
        int[] slicerColumnIndices = { 0, 1, 2 }; // adjust as needed

        // Layout parameters
        int leftPixel = 20;               // left margin for all slicers
        int startTopPixel = 20;           // top margin for the first slicer
        int slicerHeightPixel = 150;      // uniform height for each slicer
        int slicerWidthPixel = 200;       // uniform width for each slicer
        int verticalSpacingPixel = 10;    // space between consecutive slicers

        // Add a slicer for each specified column and align them vertically
        for (int i = 0; i < slicerColumnIndices.Length; i++)
        {
            int colIdx = slicerColumnIndices[i];

            // Add slicer using the overload that accepts ListObject, ListColumn, row, column.
            // Row and column here are only placeholders; we will reposition the slicer later.
            int slicerIndex = worksheet.Slicers.Add(table, table.ListColumns[colIdx], 0, 0);
            Slicer slicer = worksheet.Slicers[slicerIndex];

            // Set size
            slicer.HeightPixel = slicerHeightPixel;
            slicer.WidthPixel = slicerWidthPixel;

            // Position slicer vertically with equal spacing
            slicer.TopPixel = startTopPixel + i * (slicerHeightPixel + verticalSpacingPixel);
            slicer.LeftPixel = leftPixel;
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}