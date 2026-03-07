using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklineTypeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Define the location where the sparkline will be placed (cell E1)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group with an initial type (Column)
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Column, "A1:D1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Change the sparkline type to Line
            group.Type = SparklineType.Line;

            // Optionally, you could set it to Win/Loss (Stacked) like this:
            // group.Type = SparklineType.Stacked;

            // Add the sparkline to the group (the data range is the same as used above)
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

            // Save the workbook to an XLSX file
            workbook.Save("SparklineTypeDemo.xlsx");
        }
    }
}