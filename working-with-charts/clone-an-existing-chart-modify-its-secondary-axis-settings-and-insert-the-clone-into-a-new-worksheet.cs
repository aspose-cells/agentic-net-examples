using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartCloneDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sourceSheet = workbook.Worksheets[0];
            sourceSheet.Name = "Source";

            // Populate sample data
            sourceSheet.Cells["A1"].PutValue("Category");
            sourceSheet.Cells["A2"].PutValue("A");
            sourceSheet.Cells["A3"].PutValue("B");
            sourceSheet.Cells["A4"].PutValue("C");

            sourceSheet.Cells["B1"].PutValue("Series 1");
            sourceSheet.Cells["B2"].PutValue(100);
            sourceSheet.Cells["B3"].PutValue(200);
            sourceSheet.Cells["B4"].PutValue(300);

            sourceSheet.Cells["C1"].PutValue("Series 2");
            sourceSheet.Cells["C2"].PutValue(5000);
            sourceSheet.Cells["C3"].PutValue(3000);
            sourceSheet.Cells["C4"].PutValue(1000);

            // Add original chart
            int chartIdx = sourceSheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart originalChart = sourceSheet.Charts[chartIdx];
            originalChart.NSeries.Add("B2:B4", true);
            originalChart.NSeries.Add("C2:C4", true);
            originalChart.NSeries.CategoryData = "A2:A4";

            // Plot second series on secondary Y axis
            originalChart.NSeries[1].PlotOnSecondAxis = true;

            // Configure secondary axis for the original chart (optional, just to show it works)
            Axis secAxis = originalChart.SecondValueAxis;
            secAxis.Title.Text = "Original Secondary Axis";
            secAxis.MinValue = 0;
            secAxis.MaxValue = 6000;
            secAxis.MajorUnit = 1000;

            // Clone the worksheet (which also clones the chart)
            int clonedSheetIdx = workbook.Worksheets.AddCopy(sourceSheet.Name);
            Worksheet clonedSheet = workbook.Worksheets[clonedSheetIdx];
            clonedSheet.Name = "Cloned";

            // Access the cloned chart (same index as in source sheet)
            Chart clonedChart = clonedSheet.Charts[chartIdx];

            // Modify secondary axis settings of the cloned chart
            Axis clonedSecAxis = clonedChart.SecondValueAxis;
            clonedSecAxis.Title.Text = "Cloned Secondary Axis";
            clonedSecAxis.MinValue = 500;   // new minimum
            clonedSecAxis.MaxValue = 5500;  // new maximum
            clonedSecAxis.MajorUnit = 500;  // new major unit

            // Save the workbook
            workbook.Save("ChartCloneWithModifiedSecondaryAxis.xlsx");
        }
    }
}