using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineSeriesColorExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["A2"].PutValue(8);
            sheet.Cells["A3"].PutValue(3);

            // Define the data range for the sparkline
            string dataRange = "A1:A3";

            // Define where the sparkline will be placed (column B, rows 1-3)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 2,
                StartColumn = 1,
                EndColumn = 1
            };

            // Add a sparkline group of type Column
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Column, dataRange, false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Create a CellsColor instance and set its .Color property
            CellsColor seriesColor = workbook.CreateCellsColor();
            seriesColor.Color = Color.Blue; // Desired series color

            // Assign the series color to the sparkline group
            group.SeriesColor = seriesColor;

            // Save the workbook to an XLSX file
            workbook.Save("SparklineSeriesColorDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}