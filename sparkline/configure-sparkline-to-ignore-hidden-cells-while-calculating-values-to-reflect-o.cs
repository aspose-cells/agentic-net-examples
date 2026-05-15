using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineIgnoreHiddenDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in row 1
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Hide column B (index 1) – this column will be ignored by the sparkline
            sheet.Cells.Columns[1].IsHidden = true;

            // Define the location where the sparkline will be placed (cell E1)
            CellArea sparklineArea = new CellArea
            {
                StartColumn = 4, // Column E
                EndColumn = 4,
                StartRow = 0,    // Row 1
                EndRow = 0
            };

            // Add a sparkline group that uses the data range A1:D1
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, sparklineArea);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add the sparkline to the group (the same range is used for the sparkline)
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

            // Configure the sparkline to ignore hidden cells
            // Setting DisplayHidden to false ensures hidden rows/columns are not considered
            group.DisplayHidden = false;

            // Optional: set a series color for better visibility
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Blue;
            group.SeriesColor = seriesColor;

            // Save the workbook
            workbook.Save("SparklineIgnoreHidden.xlsx");
        }
    }
}