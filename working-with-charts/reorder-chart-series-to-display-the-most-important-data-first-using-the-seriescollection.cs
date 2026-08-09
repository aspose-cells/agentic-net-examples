// Title: Reorder Excel chart series with Aspose.Cells .NET using SeriesCollection.SwapSeries
// Description: C# example that creates a workbook, adds a column chart, computes the sum of each series directly from worksheet cells, determines a descending order by total value, and reorders the chart series with SeriesCollection.SwapSeries before saving the file.
// Keywords: Aspose.Cells | SeriesCollection | SwapSeries | C# chart series reorder | Excel chart series sorting | programmatic chart manipulation | .NET Excel automation | column chart series priority
// Common Searches: how to change series order in Aspose.Cells chart | swap series positions programmatically Aspose.Cells | sort Excel chart series by total value C# | reorder column chart series using Aspose.Cells .NET | SeriesCollection SwapSeries example
// Developer Intent: Arrange chart series so the series with the highest aggregate value appears first.
// Use Cases: Display the top‑selling product first in a sales column chart. | Show the most significant revenue stream ahead of other metrics in a financial dashboard. | Automatically prioritize KPI series based on their summed performance for executive reports.
// AI Prompts: Write C# code with Aspose.Cells that reorders chart series by descending sum of their values. | Explain the mechanics of SeriesCollection.SwapSeries and suggest an alternative using RemoveSeries and AddSeries. | Generate a dynamic solution that discovers series ranges without hard‑coded column indexes and sorts them.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesReorder
{
    // C# example that creates a workbook, adds a column chart, computes the sum of each series directly from worksheet cells, determines a descending order by total value, and reorders the chart series with SeriesCollection.SwapSeries before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for three series (columns B, C, D) with categories in column A
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            // Series 1 values (column B)
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["B5"].PutValue(40);

            // Series 2 values (column C)
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(25);
            sheet.Cells["C3"].PutValue(15);
            sheet.Cells["C4"].PutValue(35);
            sheet.Cells["C5"].PutValue(20);

            // Series 3 values (column D)
            sheet.Cells["D1"].PutValue("Series 3");
            sheet.Cells["D2"].PutValue(5);
            sheet.Cells["D3"].PutValue(10);
            sheet.Cells["D4"].PutValue(15);
            sheet.Cells["D5"].PutValue(20);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add all three series to the chart at once (range B2:D5, true = column-wise)
            chart.NSeries.Add("B2:D5", true);
            // Set category (X) axis data
            chart.NSeries.CategoryData = "A2:A5";

            // Get the series collection
            SeriesCollection seriesColl = chart.NSeries;

            // Determine importance of each series (here we use the sum of its values)
            var seriesInfo = new List<(int Index, double Sum)>();
            for (int i = 0; i < seriesColl.Count; i++)
            {
                double sum = 0;
                // Values are stored in the worksheet cells; retrieve them directly
                // Assuming each series occupies a contiguous column starting from B
                // Column index: B = 1, C = 2, D = 3 (zero‑based)
                int columnOffset = 1 + i; // B=1, C=2, D=3
                for (int row = 1; row <= 4; row++) // rows 2‑5 (zero‑based index 1‑4)
                {
                    var cell = sheet.Cells[row, columnOffset];
                    if (cell.Type == CellValueType.IsNumeric)
                        sum += cell.DoubleValue;
                }
                seriesInfo.Add((i, sum));
            }

            // Sort series indices by descending sum (most important first)
            var desiredOrder = seriesInfo.OrderByDescending(s => s.Sum).Select(s => s.Index).ToList();

            // Reorder the series using SwapSeries (simple bubble‑sort approach)
            // The goal is to transform the current order [0,1,2,...] into desiredOrder
            for (int targetPos = 0; targetPos < desiredOrder.Count; targetPos++)
            {
                int currentIdx = desiredOrder[targetPos];
                // Find where this series currently resides
                int currentPos = -1;
                for (int j = 0; j < seriesColl.Count; j++)
                {
                    if (j == currentIdx)
                    {
                        currentPos = j;
                        break;
                    }
                }
                // If the series is not at the target position, swap it forward
                while (currentPos > targetPos)
                {
                    seriesColl.SwapSeries(currentPos - 1, currentPos);
                    // Update the tracking list because indices have changed
                    for (int k = 0; k < desiredOrder.Count; k++)
                    {
                        if (desiredOrder[k] == currentPos) desiredOrder[k] = currentPos - 1;
                        else if (desiredOrder[k] == currentPos - 1) desiredOrder[k] = currentPos;
                    }
                    currentPos--;
                }
            }

            // Save the workbook
            workbook.Save("ReorderedSeriesChart.xlsx");
        }
    }
}
