// Title: How to reorder series in an Aspose.Cells column chart with SeriesCollection.Move using C#
// Common Searches: aspnet cells c# move series to top of column chart | change series order in Aspose.Cells chart programmatically | reorder chart series priority Aspose.Cells column chart C# example | using SeriesCollection.Move to set high importance series first in Excel chart | Aspose.Cells reorder series based on series name C#
// Tags: SeriesCollection.Move Aspose.Cells | reorder chart series C# | column chart series order Aspose.Cells | change series order Excel workbook C# | Aspose.Cells series priority handling

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ReorderChartSeriesDemo
{
    // The sample creates a workbook, adds Low, Medium, and High data series to a column chart, and then uses SeriesCollection.Move to rearrange the series so that the High series appears first, followed by Medium and Low. It demonstrates building a name‑based index list, moving each series to its target position, and optionally shows the obsolete ChangeSeriesOrder method. The workbook is saved as ReorderedSeriesChart.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Populate sample data for three series (Low, Medium, High)
            // ------------------------------------------------------------
            // Category labels
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            // Low importance series values (column B)
            sheet.Cells["B2"].PutValue(20);
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["B4"].PutValue(25);
            sheet.Cells["B5"].PutValue(15);

            // Medium importance series values (column C)
            sheet.Cells["C2"].PutValue(40);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(35);
            sheet.Cells["C5"].PutValue(50);

            // High importance series values (column D)
            sheet.Cells["D2"].PutValue(70);
            sheet.Cells["D3"].PutValue(80);
            sheet.Cells["D4"].PutValue(65);
            sheet.Cells["D5"].PutValue(90);

            // ------------------------------------------------------------
            // Add a column chart
            // ------------------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add three series to the chart (initial order: Low, Medium, High)
            chart.NSeries.Add("B2:B5", true); // Low
            chart.NSeries[0].Name = "Low";

            chart.NSeries.Add("C2:C5", true); // Medium
            chart.NSeries[1].Name = "Medium";

            chart.NSeries.Add("D2:D5", true); // High
            chart.NSeries[2].Name = "High";

            // Set category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A5";

            // ------------------------------------------------------------
            // Reorder series so that the most important data (High) appears first
            // ------------------------------------------------------------
            SeriesCollection seriesCollection = chart.NSeries;

            // Build a list of (currentIndex, seriesName)
            var seriesInfo = new List<(int Index, string Name)>();
            for (int i = 0; i < seriesCollection.Count; i++)
            {
                seriesInfo.Add((i, seriesCollection[i].Name));
            }

            // Desired order: High, Medium, Low
            var desiredOrder = new List<string> { "High", "Medium", "Low" };

            // Reorder using Series.Move (preferred) – move each series to its target position
            foreach (var targetName in desiredOrder)
            {
                // Find the series that currently has this name
                int currentIdx = seriesInfo.FindIndex(s => s.Name.Equals(targetName, StringComparison.OrdinalIgnoreCase));
                if (currentIdx == -1)
                    continue; // safety check

                // Desired position in the chart (0‑based)
                int targetPosition = desiredOrder.IndexOf(targetName);

                // If the series is already at the correct position, skip
                if (currentIdx == targetPosition)
                    continue;

                // Move the series to the target position
                seriesCollection[currentIdx].Move(targetPosition);

                // After moving, update the indices in seriesInfo to reflect the new order
                // The moved series now resides at targetPosition; shift others accordingly
                var movedSeries = seriesInfo[currentIdx];
                seriesInfo.RemoveAt(currentIdx);
                seriesInfo.Insert(targetPosition, movedSeries);
            }

            // ------------------------------------------------------------
            // (Optional) Demonstrate the obsolete ChangeSeriesOrder method
            // ------------------------------------------------------------
#pragma warning disable 618 // Suppress obsolete warning
            // Swap series 0 and 2 using the obsolete method (just for illustration)
            // seriesCollection.ChangeSeriesOrder(0, 2);
#pragma warning restore 618

            // ------------------------------------------------------------
            // Save the workbook
            // ------------------------------------------------------------
            string outputPath = "ReorderedSeriesChart.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'. Series have been reordered: High, Medium, Low.");
        }
    }
}
