// Title: Generate an Excel combo chart with column, line, and area series on primary and secondary axes using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create combo chart with column, line and area series | How to plot a line series on the secondary axis in Aspose.Cells | Add multiple chart types to a single chart with Aspose.Cells .NET | Set secondary value axis for a series in an Aspose.Cells workbook | Create a combo chart with different series types using Aspose.Cells C#
// Tags: combo chart multiple series Aspose.Cells | column line area chart secondary axis | plot series on second axis Aspose.Cells | Aspose.Cells chart type customization C# | Excel combo chart generation .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ComboChartExample
{
    // The example creates a new workbook, fills cells A2:D5 with quarterly data, adds a column‑based combo chart, and defines three series: a column series on the primary axis, and line and area series plotted on the secondary axis. It configures axis titles, chart title, legend position, and saves the file as ComboChart_ColumnLineArea.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Category labels
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            // Column series data
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["B5"].PutValue(200);

            // Line series data
            sheet.Cells["C2"].PutValue(80);
            sheet.Cells["C3"].PutValue(130);
            sheet.Cells["C4"].PutValue(170);
            sheet.Cells["C5"].PutValue(210);

            // Area series data
            sheet.Cells["D2"].PutValue(60);
            sheet.Cells["D3"].PutValue(110);
            sheet.Cells["D4"].PutValue(140);
            sheet.Cells["D5"].PutValue(190);

            // Add a combo chart (base type Column) to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the category (X) axis data
            chart.NSeries.CategoryData = "A2:A5";

            // ----- Series 1: Column (Primary Axis) -----
            int colSeriesIdx = chart.NSeries.Add("B2:B5", true);
            Series colSeries = chart.NSeries[colSeriesIdx];
            colSeries.Type = ChartType.Column;          // Explicitly set type (optional, default is Column)
            colSeries.Name = "Sales (Column)";
            colSeries.PlotOnSecondAxis = false;         // Primary axis

            // ----- Series 2: Line (Secondary Axis) -----
            int lineSeriesIdx = chart.NSeries.Add("C2:C5", true);
            Series lineSeries = chart.NSeries[lineSeriesIdx];
            lineSeries.Type = ChartType.Line;
            lineSeries.Name = "Profit (Line)";
            lineSeries.PlotOnSecondAxis = true;         // Secondary axis

            // ----- Series 3: Area (Secondary Axis) -----
            int areaSeriesIdx = chart.NSeries.Add("D2:D5", true);
            Series areaSeries = chart.NSeries[areaSeriesIdx];
            areaSeries.Type = ChartType.Area;
            areaSeries.Name = "Cost (Area)";
            areaSeries.PlotOnSecondAxis = true;         // Secondary axis (different visual axis from column)

            // Optional: customize axes titles
            chart.CategoryAxis.Title.Text = "Quarter";
            chart.ValueAxis.Title.Text = "Primary Axis (Units)";
            chart.SecondValueAxis.Title.Text = "Secondary Axis (Units)";

            // Chart title and legend
            chart.Title.Text = "Combo Chart: Column, Line, and Area";
            chart.ShowLegend = true;
            chart.Legend.Position = LegendPositionType.Bottom;

            // Save the workbook
            string outputPath = "ComboChart_ColumnLineArea.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
