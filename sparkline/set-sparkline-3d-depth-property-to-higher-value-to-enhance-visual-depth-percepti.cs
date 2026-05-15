using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsExamples
{
    class Sparkline3DDepthDemo
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(90);
            sheet.Cells["C3"].PutValue(110);
            sheet.Cells["C4"].PutValue(130);

            // Add a 3‑D column chart and increase its depth
            int chartIdx = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.DepthPercent = 300; // 300% depth

            // Define the location range for sparklines (E2:E4)
            CellArea sparklineArea = new CellArea
            {
                StartColumn = 4, // E column (0‑based)
                EndColumn = 4,
                StartRow = 1,    // row 2 (0‑based)
                EndRow = 3       // row 4 (0‑based)
            };

            // Add a sparkline group (line type)
            int sparklineGroupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "B2:C4", false, sparklineArea);
            SparklineGroup sparklineGroup = sheet.SparklineGroups[sparklineGroupIdx];

            // Optional visual tweaks
            sparklineGroup.LineWeight = 1.5;
            CellsColor sparkColor = workbook.CreateCellsColor();
            sparkColor.Color = Color.Orange;
            sparklineGroup.SeriesColor = sparkColor;

            // Save the workbook
            workbook.Save("Sparkline3DDepthDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}