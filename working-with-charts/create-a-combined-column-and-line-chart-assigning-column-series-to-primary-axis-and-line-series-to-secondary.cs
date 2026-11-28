// Title: Create a combined column and line chart with primary and secondary axes in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create combined column and line chart with secondary axis | How to add a line series on the secondary axis in an Aspose.Cells mixed chart | Set column series to primary axis and line series to secondary axis using Aspose.Cells | C# Aspose.Cells mixed chart axis configuration example | Generate a column‑line chart with separate axes in Aspose.Cells .NET
// Tags: column line mixed chart Aspose.Cells | secondary axis line series Aspose.Cells | primary axis column series Aspose.Cells | chart axis titles Aspose.Cells | save chart as XLSX Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace CombinedColumnLineChartDemo
{
    // The example creates a new workbook, populates it with monthly sales and growth data, adds a mixed column‑line chart, assigns the column series to the primary axis and the line series to the secondary axis, configures axis and chart titles, and saves the workbook as an XLSX file using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Populate sample data
                // -------------------------------------------------
                // Category (X‑axis) values
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");
                sheet.Cells["A6"].PutValue("May");

                // Column series data (Primary axis)
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(170);
                sheet.Cells["B5"].PutValue(130);
                sheet.Cells["B6"].PutValue(190);

                // Line series data (Secondary axis)
                sheet.Cells["C1"].PutValue("Growth %");
                sheet.Cells["C2"].PutValue(5);
                sheet.Cells["C3"].PutValue(7);
                sheet.Cells["C4"].PutValue(6);
                sheet.Cells["C5"].PutValue(8);
                sheet.Cells["C6"].PutValue(9);

                // -------------------------------------------------
                // Add a combined chart (Column + Line)
                // -------------------------------------------------
                // Add a column chart – this will be the base chart type
                int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart title
                chart.Title.Text = "Monthly Sales and Growth";

                // -------------------------------------------------
                // Add the column series (primary axis)
                // -------------------------------------------------
                int colSeriesIdx = chart.NSeries.Add("B2:B6", true);
                Series colSeries = chart.NSeries[colSeriesIdx];
                colSeries.Name = "Sales";
                colSeries.Type = ChartType.Column;               // Explicitly set column type
                colSeries.PlotOnSecondAxis = false;              // Primary axis (default)

                // -------------------------------------------------
                // Add the line series (secondary axis)
                // -------------------------------------------------
                int lineSeriesIdx = chart.NSeries.Add("C2:C6", true);
                Series lineSeries = chart.NSeries[lineSeriesIdx];
                lineSeries.Name = "Growth %";
                lineSeries.Type = ChartType.Line;                // Change to line chart
                lineSeries.PlotOnSecondAxis = true;              // Plot on secondary axis

                // Optional: set primary axis title
                chart.ValueAxis.Title.Text = "Sales (Units)";
                chart.SecondValueAxis.Title.Text = "Growth (%)";

                // Optional: set category axis title
                chart.CategoryAxis.Title.Text = "Month";

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                string outputPath = "CombinedColumnLineChart.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the chart:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
