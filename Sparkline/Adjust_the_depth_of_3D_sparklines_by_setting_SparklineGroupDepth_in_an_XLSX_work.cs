using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AdjustSparklineDepthDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (A1:A5)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(3);
            sheet.Cells["A3"].PutValue(7);
            sheet.Cells["A4"].PutValue(2);
            sheet.Cells["A5"].PutValue(9);

            // Define the location where the sparkline will be placed (B1:B5)
            CellArea location = CellArea.CreateCellArea("B1", "B5");

            // Add a sparkline group of type Column (3‑D effect) with the data range and location
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Column, "A1:A5", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // No Depth property exists; the Column sparkline inherently displays a 3‑D effect.

            // Save the workbook to an XLSX file
            workbook.Save("SparklineDepthAdjusted.xlsx");
        }
    }
}