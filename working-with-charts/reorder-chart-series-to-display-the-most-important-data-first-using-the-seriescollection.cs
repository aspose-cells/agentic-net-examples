// Title: How to reorder chart series in Aspose.Cells C# by moving the most important series to the front using SeriesCollection.SwapSeries
// AI Prompts: Generate C# code that uses Aspose.Cells to swap the positions of two series in a column chart so the highest‑priority series appears first. | Show how to reorder multiple series in an Excel column chart created with Aspose.Cells by calling SeriesCollection.SwapSeries with specific indices. | Provide a step‑by‑step C# example that creates a workbook, adds a column chart, and changes the series display order to prioritize a selected series.
// Common Searches: aspnet change chart series order in excel using aspose.cells | c# move most important series to first position in column chart | how to prioritize a series in an Aspose.Cells generated chart | swap series indices in Aspose.Cells chart programmatically | excel chart series ordering example with Aspose.Cells .NET
// Tags: Aspose.Cells SeriesCollection.SwapSeries method | column chart series priority handling | set series display order in Excel chart C# | Aspose.Cells chart series manipulation | C# Excel chart series ordering

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesReorder
{
    // The sample creates a workbook, populates three data series, adds a column chart, and then uses SeriesCollection.SwapSeries to rearrange the series so the most important one is displayed first before saving the file as ReorderedSeriesChart.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for three series
            // Category labels
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            // Series 1 (least important)
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Series 2 (medium importance)
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(40);
            sheet.Cells["C3"].PutValue(50);
            sheet.Cells["C4"].PutValue(60);

            // Series 3 (most important)
            sheet.Cells["D1"].PutValue("Series 3");
            sheet.Cells["D2"].PutValue(70);
            sheet.Cells["D3"].PutValue(80);
            sheet.Cells["D4"].PutValue(90);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Add all three series to the chart (by column)
            chart.NSeries.Add("B1:D4", true);

            // Set category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A4";

            // Access the series collection
            SeriesCollection seriesColl = chart.NSeries;

            // Original order: Series 1, Series 2, Series 3
            // Desired order: Series 3 (most important), Series 2, Series 1
            // Perform swaps using SwapSeries method
            // Move Series 3 (index 2) to index 0
            seriesColl.SwapSeries(2, 0);   // After this: Series 3, Series 1, Series 2
            // Move Series 2 (now at index 2) to index 1
            seriesColl.SwapSeries(2, 1);   // Final order: Series 3, Series 2, Series 1

            // Save the workbook with the reordered chart
            workbook.Save("ReorderedSeriesChart.xlsx");
        }
    }
}
