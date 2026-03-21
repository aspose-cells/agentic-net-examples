using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class Sparkline3DDepthDemo
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(90);
            sheet.Cells["C3"].PutValue(110);
            sheet.Cells["C4"].PutValue(130);

            // Add a 3‑D column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Modify the depth of the 3‑D chart (percentage of chart width)
            chart.DepthPercent = 150; // 150% depth

            // Save the workbook as XLSX
            workbook.Save("Sparkline3DDepthDemo.xlsx", SaveFormat.Xlsx);
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            Sparkline3DDepthDemo.Run();
        }
    }
}