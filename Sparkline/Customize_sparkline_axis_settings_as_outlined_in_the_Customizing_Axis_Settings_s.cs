using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineAxisCustomization
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
            sheet.Cells["B1"].PutValue(-2);
            sheet.Cells["C1"].PutValue(3);
            sheet.Cells["D1"].PutValue(-1);

            // Define the location where the sparkline will be placed (E1)
            CellArea sparklineLocation = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4, // Column E (0‑based index)
                EndColumn = 4
            };

            // Add a sparkline group with the data range A1:D1
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, sparklineLocation);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Show the horizontal axis and set its color
            group.ShowHorizontalAxis = true;
            CellsColor hAxisColor = workbook.CreateCellsColor();
            hAxisColor.Color = Color.Gray;
            group.HorizontalAxisColor = hAxisColor;

            // Set a date range for the horizontal axis (optional, demonstrates the property)
            // Here we use dummy dates in the first row
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["B2"].PutValue(new DateTime(2023, 2, 1));
            sheet.Cells["C2"].PutValue(new DateTime(2023, 3, 1));
            sheet.Cells["D2"].PutValue(new DateTime(2023, 4, 1));
            group.HorizontalAxisDateRange = "A2:D2";
            group.ShowHorizontalAxis = true;

            // Configure vertical axis settings
            // Use the same max/min for all sparklines in the group
            group.VerticalAxisMaxValueType = SparklineAxisMinMaxType.Group;
            group.VerticalAxisMaxValue = 6.0; // Custom max value
            group.VerticalAxisMinValueType = SparklineAxisMinMaxType.Group;
            group.VerticalAxisMinValue = -3.0; // Custom min value

            // Save the workbook
            workbook.Save("SparklineCustomAxis.xlsx");
        }
    }
}