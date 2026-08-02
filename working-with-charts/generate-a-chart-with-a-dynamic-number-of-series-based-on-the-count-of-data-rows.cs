using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace DynamicSeriesChartDemo
{
    class Program
    {
        // Helper to convert zero‑based column index to Excel column name (A, B, …, AA, AB, …)
        static string GetColumnName(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string name = "";
            while (index >= 0)
            {
                name = letters[index % 26] + name;
                index = index / 26 - 1;
            }
            return name;
        }

        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------- Sample data preparation --------------------
            int dataRows = 6;          // Number of data rows (excluding header)
            int seriesCount = 4;       // Number of series (columns) – this can be dynamic

            // Header row
            sheet.Cells["A1"].PutValue("Category");
            for (int s = 0; s < seriesCount; s++)
            {
                // Series names in B1, C1, …
                sheet.Cells[0, s + 1].PutValue($"Series {s + 1}");
            }

            // Fill categories and random values
            for (int r = 0; r < dataRows; r++)
            {
                // Category labels in column A (A2, A3, …)
                sheet.Cells[r + 1, 0].PutValue($"Cat {r + 1}");

                // Values for each series column
                for (int s = 0; s < seriesCount; s++)
                {
                    // Example value: (r+1)*(s+1)*10
                    sheet.Cells[r + 1, s + 1].PutValue((r + 1) * (s + 1) * 10);
                }
            }

            // -------------------- Chart creation --------------------
            // Add a column chart positioned on the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set category (X‑axis) data for all series
            chart.NSeries.CategoryData = $"=Sheet1!$A$2:$A${dataRows + 1}";

            // Dynamically add a series for each data column
            for (int s = 0; s < seriesCount; s++)
            {
                // Column letter for the current series (B, C, D, …)
                string colLetter = GetColumnName(s + 1);
                // Data range for the series (e.g., =Sheet1!$B$2:$B$7)
                string dataRange = $"=Sheet1!${colLetter}$2:${colLetter}${dataRows + 1}";
                // Add the series; true = plot by column (vertical)
                chart.NSeries.Add(dataRange, true);
            }

            // Optional: give the chart a title
            chart.Title.Text = "Dynamic Series Chart";

            // -------------------- Save the workbook --------------------
            workbook.Save("DynamicSeriesChart.xlsx");
        }
    }
}