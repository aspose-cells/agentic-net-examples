using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace Adjust3DSparklineDepth
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for sparklines and the 3‑D chart
            // Data for sparklines (row 1)
            sheet.Cells["A1"].PutValue(5);
            sheet.Cells["B1"].PutValue(3);
            sheet.Cells["C1"].PutValue(7);
            sheet.Cells["D1"].PutValue(2);

            // Data for the 3‑D chart (columns B and C)
            sheet.Cells["A2"].PutValue("Category");
            sheet.Cells["A3"].PutValue("Q1");
            sheet.Cells["A4"].PutValue("Q2");
            sheet.Cells["A5"].PutValue("Q3");

            sheet.Cells["B2"].PutValue("Series 1");
            sheet.Cells["B3"].PutValue(120);
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["B5"].PutValue(180);

            sheet.Cells["C2"].PutValue("Series 2");
            sheet.Cells["C3"].PutValue(90);
            sheet.Cells["C4"].PutValue(110);
            sheet.Cells["C5"].PutValue(130);

            // -------------------------------------------------
            // Add a sparkline group (2‑D, as sparklines are 2‑D)
            // -------------------------------------------------
            CellArea sparklineArea = CellArea.CreateCellArea("E1", "E1"); // place sparkline in column E, row 1
            int sparklineGroupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, sparklineArea);
            SparklineGroup sparklineGroup = sheet.SparklineGroups[sparklineGroupIdx];
            // (Optional) customize sparkline appearance
            sparklineGroup.ShowHighPoint = true;
            sparklineGroup.ShowLowPoint = true;

            // -------------------------------------------------
            // Add a 3‑D column chart and adjust its depth
            // -------------------------------------------------
            // Add the chart (positioned starting at row 7, column 0)
            int chartIdx = sheet.Charts.Add(ChartType.Column3D, 7, 0, 20, 8);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the chart
            chart.NSeries.Add("B3:C5", true);          // values
            chart.NSeries.CategoryData = "A3:A5";      // categories

            // Adjust the depth of the 3‑D chart.
            // DepthPercent represents the depth as a percentage of the chart width (20‑2000%).
            chart.DepthPercent = 150; // 150 % depth – this is the property that controls 3‑D depth.

            // Save the workbook in XLSX format
            workbook.Save("Adjusted3DSparklineDepth.xlsx", SaveFormat.Xlsx);
        }
    }
}