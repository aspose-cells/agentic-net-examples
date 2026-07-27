using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartFillExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
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

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure the chart area to use a solid fill with 80% opacity (20% transparency)
            chart.ChartArea.Area.FillFormat.FillType = FillType.Solid;                     // Use solid fill
            chart.ChartArea.Area.FillFormat.SolidFill.Color = Color.LightGray;           // Choose a subtle color
            chart.ChartArea.Area.FillFormat.SolidFill.Transparency = 0.2;                // 20% transparent = 80% opacity

            // Save the workbook
            workbook.Save("ChartWithSolidFill.xlsx");
        }
    }
}