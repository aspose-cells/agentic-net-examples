using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendResizeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.SetChartDataRange("A1:B4", true);

            // -----------------------------------------------------------------
            // Step 1: Set legend to a fixed size (disable automatic sizing)
            // -----------------------------------------------------------------
            Legend legend = chart.Legend;
            legend.IsAutomaticSize = false;          // Fixed size mode
            legend.WidthPixel = 200;                  // Fixed width in pixels
            legend.HeightPixel = 60;                  // Fixed height in pixels
            legend.Position = LegendPositionType.Bottom;
            legend.Font.Size = 12;
            legend.Font.IsBold = true;

            // -----------------------------------------------------------------
            // Step 2: Simulate a change that would affect legend content
            // (e.g., add a longer series name)
            // -----------------------------------------------------------------
            // Add a second series with a long name to force legend to need more space
            sheet.Cells["C1"].PutValue("Long Series Name Example");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries[1].Name = "Long Series Name Example";

            // -----------------------------------------------------------------
            // Step 3: Adjust legend size based on the new content
            // -----------------------------------------------------------------
            // Enable automatic sizing, recalculate the chart, then capture the
            // automatically determined size and apply it back as a fixed size.
            legend.IsAutomaticSize = true;            // Let Excel compute needed size
            chart.Calculate();                       // Force layout recalculation

            // Retrieve the automatically calculated size (in pixels)
            int autoWidth = legend.WidthPixel;
            int autoHeight = legend.HeightPixel;

            // Apply the calculated size as a fixed size again
            legend.IsAutomaticSize = false;
            legend.WidthPixel = autoWidth;
            legend.HeightPixel = autoHeight;

            // -----------------------------------------------------------------
            // Save the workbook
            // -----------------------------------------------------------------
            workbook.Save("ChartLegendAutoResizeDemo.xlsx");
        }
    }
}