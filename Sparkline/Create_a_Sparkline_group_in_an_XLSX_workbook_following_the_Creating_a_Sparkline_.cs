using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineGroupDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (A1:D1)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(3);
            sheet.Cells["C1"].PutValue(8);
            sheet.Cells["D1"].PutValue(2);

            // Define the location where the sparkline will be placed (cell E1)
            CellArea location = new CellArea
            {
                StartRow = 0,   // Row 0 (A1 is row 0)
                EndRow = 0,
                StartColumn = 4, // Column 4 corresponds to column E
                EndColumn = 4
            };

            // Add a sparkline group of type Line, using the data range A1:D1,
            // plotting horizontally (isVertical = false), and placing the sparkline in E1
            int groupIndex = sheet.SparklineGroups.Add(
                SparklineType.Line,
                "A1:D1",
                false,
                location);

            // Retrieve the created SparklineGroup
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Add a sparkline to the group.
            // The data range is the same as above; row and column specify the location (E1).
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

            // Optional: customize appearance (e.g., set series color and show high/low points)
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Orange;
            group.SeriesColor = seriesColor;

            group.ShowHighPoint = true;
            group.ShowLowPoint = true;

            CellsColor highColor = workbook.CreateCellsColor();
            highColor.Color = Color.Green;
            group.HighPointColor = highColor;

            CellsColor lowColor = workbook.CreateCellsColor();
            lowColor.Color = Color.Red;
            group.LowPointColor = lowColor;

            // Save the workbook as an XLSX file
            workbook.Save("SparklineGroupDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}