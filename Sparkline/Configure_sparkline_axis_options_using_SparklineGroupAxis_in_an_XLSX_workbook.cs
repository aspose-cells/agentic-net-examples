using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineAxisConfigurationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (values)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(2);
            sheet.Cells["C1"].PutValue(1);
            sheet.Cells["D1"].PutValue(3);

            // Populate date values for the horizontal axis (optional)
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["B2"].PutValue(new DateTime(2023, 2, 1));
            sheet.Cells["C2"].PutValue(new DateTime(2023, 3, 1));
            sheet.Cells["D2"].PutValue(new DateTime(2023, 4, 1));

            // Define where the sparkline will be placed (cell E1)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group: line type, data range A1:D1, horizontal orientation, location defined above
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Add the sparkline to the group (the Add method also creates the sparkline item)
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

            // ----- Configure horizontal axis options -----
            // Show the horizontal axis (useful when data crosses zero)
            group.ShowHorizontalAxis = true;

            // Set the color of the horizontal axis line
            CellsColor axisColor = workbook.CreateCellsColor();
            axisColor.Color = Color.Gray;
            group.HorizontalAxisColor = axisColor;

            // Associate a date range with the horizontal axis (optional)
            group.HorizontalAxisDateRange = "A2:D2";

            // ----- Configure vertical axis options -----
            // Use the same max/min values for all sparklines in the group
            group.VerticalAxisMaxValueType = SparklineAxisMinMaxType.Group;
            group.VerticalAxisMinValueType = SparklineAxisMinMaxType.Group;

            // Set custom max/min values
            group.VerticalAxisMaxValue = 10.0;
            group.VerticalAxisMinValue = 0.0;

            // Save the workbook to an XLSX file
            workbook.Save("SparklineAxisConfigured.xlsx");
        }
    }
}