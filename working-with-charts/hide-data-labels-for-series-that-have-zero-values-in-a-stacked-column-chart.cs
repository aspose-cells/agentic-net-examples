// Title: How to hide data labels for zero‑value points in a stacked column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# with Aspose.Cells to create a stacked column chart and automatically hide data labels for any data point whose value is zero. | Provide a code snippet that iterates over each series in an Aspose.Cells chart and sets DataLabels.ShowValue to false for zero‑valued points.
// Common Searches: how to suppress zero-value labels in a stacked column chart with Aspose.Cells | C# Aspose.Cells conditional visibility of chart data labels | remove data labels for zero points in Excel stacked column chart programmatically | iterate over series points to hide labels using Aspose.Cells .NET | Aspose.Cells chart label customization based on cell values
// Tags: Aspose.Cells hide zero data labels | stacked column chart label control C# | conditional chart data label Aspose.Cells | chart series point iteration Aspose.Cells | Excel chart data label suppression .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a workbook, fills it with sample data containing zeros, adds a stacked column chart, enables data labels for all series, then loops through each series point to hide the label when the underlying cell value is zero, and finally saves the workbook as an XLSX file.
    public class HideZeroValueDataLabelsInStackedColumnChart
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with some zero values
            // Category column
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            // Series 1 values
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(0);   // zero value
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(20);

            // Series 2 values
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(5);
            sheet.Cells["C3"].PutValue(15);
            sheet.Cells["C4"].PutValue(0);   // zero value
            sheet.Cells["C5"].PutValue(25);

            // Add a stacked column chart
            int chartIndex = sheet.Charts.Add(ChartType.ColumnStacked, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart (both series)
            chart.NSeries.Add("B2:C5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Enable data labels for each series
            foreach (Series series in chart.NSeries)
            {
                series.DataLabels.ShowValue = true; // show values by default
            }

            // Hide data labels for points whose value is zero
            // The data for each series starts in column B (index 1) and C (index 2)
            for (int s = 0; s < chart.NSeries.Count; s++)
            {
                Series series = chart.NSeries[s];
                int dataColumn = 1 + s; // B = 1, C = 2 (zero‑based column index)

                for (int i = 0; i < series.Points.Count; i++)
                {
                    // Data rows start at row 2 (zero‑based index 1)
                    int dataRow = 1 + i;
                    object cellValue = sheet.Cells[dataRow, dataColumn].Value;

                    double pointValue = 0;
                    if (cellValue != null && double.TryParse(cellValue.ToString(), out double parsed))
                    {
                        pointValue = parsed;
                    }

                    if (Math.Abs(pointValue) < double.Epsilon)
                    {
                        // Hide the data label for this specific point
                        series.Points[i].DataLabels.ShowValue = false;
                    }
                }
            }

            // Define output file path
            string outputPath = "HideZeroValueDataLabelsStackedColumnChart.xlsx";

            try
            {
                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
