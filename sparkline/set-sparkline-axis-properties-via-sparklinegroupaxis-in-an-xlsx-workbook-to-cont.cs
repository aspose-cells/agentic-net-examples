using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineAxisDemo
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

            // Populate date data for the horizontal axis (optional)
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["B2"].PutValue(new DateTime(2023, 2, 1));
            sheet.Cells["C2"].PutValue(new DateTime(2023, 3, 1));
            sheet.Cells["D2"].PutValue(new DateTime(2023, 4, 1));

            // Define where the sparkline will be placed (column E, row 1)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group (Line type) using the data range A1:D1
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Add the sparkline to the group (required when using Add with dataRange)
            group.Sparklines.Add(sheet.Name + "!A1:D1", 0, 4);

            // ----- Horizontal axis settings -----
            // Show the horizontal axis
            group.ShowHorizontalAxis = true;

            // Set the color of the horizontal axis
            CellsColor hAxisColor = workbook.CreateCellsColor();
            hAxisColor.Color = Color.DarkGray;
            group.HorizontalAxisColor = hAxisColor;

            // Associate a date range with the horizontal axis
            group.HorizontalAxisDateRange = "A2:D2";

            // ----- Vertical axis settings -----
            // Use the same max/min values for all sparklines in the group
            group.VerticalAxisMaxValueType = SparklineAxisMinMaxType.Group;
            group.VerticalAxisMinValueType = SparklineAxisMinMaxType.Group;

            // Define custom max/min values
            group.VerticalAxisMaxValue = 10.0;
            group.VerticalAxisMinValue = 0.0;

            // Save the workbook to an XLSX file
            workbook.Save("SparklineAxisSettings.xlsx");
        }
    }
}