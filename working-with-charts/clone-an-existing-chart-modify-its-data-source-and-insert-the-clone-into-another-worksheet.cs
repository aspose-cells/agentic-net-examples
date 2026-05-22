using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartCloneDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // ---------- Source worksheet with original chart ----------
            Worksheet srcSheet = workbook.Worksheets[0];
            srcSheet.Name = "Source";

            // Populate source data (A1:B4)
            srcSheet.Cells["A1"].PutValue("Category");
            srcSheet.Cells["B1"].PutValue("Value");
            srcSheet.Cells["A2"].PutValue("A");
            srcSheet.Cells["B2"].PutValue(10);
            srcSheet.Cells["A3"].PutValue("B");
            srcSheet.Cells["B3"].PutValue(20);
            srcSheet.Cells["A4"].PutValue("C");
            srcSheet.Cells["B4"].PutValue(30);

            // Add a chart to the source sheet
            int srcChartIdx = srcSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart srcChart = srcSheet.Charts[srcChartIdx];
            // Set data range for the source chart
            srcChart.NSeries.Add("B2:B4", true);
            srcChart.NSeries.CategoryData = "A2:A4";
            srcChart.Title.Text = "Source Chart";

            // ---------- Destination worksheet ----------
            Worksheet destSheet = workbook.Worksheets.Add("Destination");

            // Populate destination data (A1:B4) – different values to demonstrate the new source
            destSheet.Cells["A1"].PutValue("Category");
            destSheet.Cells["B1"].PutValue("Value");
            destSheet.Cells["A2"].PutValue("X");
            destSheet.Cells["B2"].PutValue(40);
            destSheet.Cells["A3"].PutValue("Y");
            destSheet.Cells["B3"].PutValue(50);
            destSheet.Cells["A4"].PutValue("Z");
            destSheet.Cells["B4"].PutValue(60);

            // ---------- Clone the chart ----------
            // Add a new chart to the destination sheet with the same type as the source chart
            int clonedChartIdx = destSheet.Charts.Add(srcChart.Type, 5, 0, 15, 5);
            Chart clonedChart = destSheet.Charts[clonedChartIdx];

            // Copy basic properties (title, style, etc.) from the source chart
            clonedChart.Title.Text = srcChart.Title.Text + " (Cloned)";
            clonedChart.Style = srcChart.Style;
            clonedChart.ShowLegend = srcChart.ShowLegend;

            // Modify the data source to refer to the destination sheet's data
            // Using SetChartDataRange automatically updates series and category data
            clonedChart.SetChartDataRange("Destination!A1:B4", true);

            // Save the workbook
            workbook.Save("ChartCloneDemo.xlsx");
        }
    }
}