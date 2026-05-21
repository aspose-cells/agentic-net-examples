using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace DynamicSeriesChartDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Prepare sample data:
            // Column A  -> Category (e.g., "Item 1", "Item 2", ...)
            // Columns B..N -> Values for each series
            // -------------------------------------------------
            int totalRows = 8;          // Number of data rows (excluding header)
            int totalSeries = 4;        // Number of series (columns B, C, D, E)

            // Header row
            sheet.Cells["A1"].PutValue("Category");
            for (int s = 0; s < totalSeries; s++)
            {
                sheet.Cells[0, s + 1].PutValue($"Series {s + 1}");
            }

            // Fill data rows
            for (int r = 0; r < totalRows; r++)
            {
                // Category label
                sheet.Cells[r + 1, 0].PutValue($"Item {r + 1}");

                // Values for each series
                for (int s = 0; s < totalSeries; s++)
                {
                    // Example value: (r + 1) * (s + 1) * 10
                    sheet.Cells[r + 1, s + 1].PutValue((r + 1) * (s + 1) * 10);
                }
            }

            // -------------------------------------------------
            // Add a chart to the worksheet
            // -------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the category (X‑axis) data for the chart
            // Category range: A2:A{totalRows+1}
            chart.NSeries.CategoryData = $"=Sheet1!$A$2:$A${totalRows + 1}";

            // Dynamically add a series for each data column (B, C, D, ...)
            for (int s = 0; s < totalSeries; s++)
            {
                // Column letter for the current series (B = 2, C = 3, ...)
                char colLetter = (char)('B' + s);
                // Data range for the series: e.g., B2:B9
                string dataRange = $"=Sheet1!${colLetter}$2:${colLetter}${totalRows + 1}";
                // Add the series; 'true' indicates vertical (column‑wise) data
                chart.NSeries.Add(dataRange, true);
            }

            // Optional: give the chart a title
            chart.Title.Text = "Dynamic Series Chart";

            // Save the workbook
            workbook.Save("DynamicSeriesChart.xlsx");
        }
    }
}