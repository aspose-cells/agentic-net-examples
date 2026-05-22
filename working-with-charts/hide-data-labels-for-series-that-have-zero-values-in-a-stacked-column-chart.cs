using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class HideZeroValueDataLabelsInStackedColumnChart
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (including zero values)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(0);   // Zero value – label should be hidden
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(0);   // Zero value – label should be hidden

            // Add a stacked column chart
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 6, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (both series)
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for all series
            foreach (Series ser in chart.NSeries)
            {
                ser.DataLabels.ShowValue = true;
            }

            // Hide data labels for points whose value is zero
            for (int s = 0; s < chart.NSeries.Count; s++)
            {
                Series ser = chart.NSeries[s];
                // Column index for the series (B = 1, C = 2, …) – zero‑based column index
                int seriesCol = 1 + s;

                for (int i = 0; i < ser.Points.Count; i++)
                {
                    // Row index for the point (data starts at row 2 → index 1)
                    int rowIndex = i + 1;

                    double pointValue = sheet.Cells[rowIndex, seriesCol].DoubleValue;

                    if (Math.Abs(pointValue) < double.Epsilon) // value is zero
                    {
                        // Hide the data label for this specific point
                        ser.Points[i].DataLabels.ShowValue = false;
                    }
                }
            }

            // Save the workbook
            string outputPath = "StackedColumnChart_HideZeroLabels.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}