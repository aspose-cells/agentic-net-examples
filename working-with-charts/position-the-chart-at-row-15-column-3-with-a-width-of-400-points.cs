using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartPositionExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart; initial position is arbitrary
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Move the chart so that its upper‑left corner is at row 15, column 3
            // BottomRow and RightColumn are set equal to TopRow/LeftColumn to keep the current size;
            // the size will be adjusted next.
            chart.Move(15, 3, 15, 3);

            // Set the chart width to 400 points (1 point = 1/72 inch)
            chart.ChartObject.WidthPt = 400;

            // Save the workbook
            workbook.Save("ChartPositioned.xlsx");
        }
    }
}