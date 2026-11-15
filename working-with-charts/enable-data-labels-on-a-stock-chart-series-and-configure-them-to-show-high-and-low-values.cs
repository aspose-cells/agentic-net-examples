// Title: How to add data labels showing high and low values to a Stock Open‑High‑Low‑Close chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# display high and low values on OHLC stock chart data labels | C# add data labels to StockOpenHighLowClose chart with Aspose.Cells | show OHLC values in an Excel stock chart using Aspose.Cells .NET | enable data labels for each point in a stock chart via Aspose.Cells C# example | Aspose.Cells example for configuring data labels on a stock chart series
// Tags: Aspose.Cells add data labels to stock chart | C# StockOpenHighLowClose chart data labels | OHLC chart data labels Aspose.Cells | Excel stock chart high low values C# | Aspose.Cells chart series label configuration

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, fills it with Open‑High‑Low‑Close data, inserts a StockOpenHighLowClose chart, adds the OHLC series, enables data labels by setting ShowValue to true, and saves the file as an Excel workbook.
class StockChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a stock chart (Open, High, Low, Close)
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Open");
            sheet.Cells["C1"].PutValue("High");
            sheet.Cells["D1"].PutValue("Low");
            sheet.Cells["E1"].PutValue("Close");

            sheet.Cells["A2"].PutValue("2023-01-01");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["C2"].PutValue(130);
            sheet.Cells["D2"].PutValue(115);
            sheet.Cells["E2"].PutValue(125);

            sheet.Cells["A3"].PutValue("2023-01-02");
            sheet.Cells["B3"].PutValue(126);
            sheet.Cells["C3"].PutValue(135);
            sheet.Cells["D3"].PutValue(122);
            sheet.Cells["E3"].PutValue(132);

            sheet.Cells["A4"].PutValue("2023-01-03");
            sheet.Cells["B4"].PutValue(133);
            sheet.Cells["C4"].PutValue(140);
            sheet.Cells["D4"].PutValue(130);
            sheet.Cells["E4"].PutValue(138);

            sheet.Cells["A5"].PutValue("2023-01-04");
            sheet.Cells["B5"].PutValue(139);
            sheet.Cells["C5"].PutValue(145);
            sheet.Cells["D5"].PutValue(135);
            sheet.Cells["E5"].PutValue(142);

            // Add a Stock chart (Open-High-Low-Close) to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.StockOpenHighLowClose, 7, 0, 25, 10);
            Chart stockChart = sheet.Charts[chartIndex];

            // Add the series: Open, High, Low, Close (order required by StockOpenHighLowClose)
            stockChart.NSeries.Add("B2:B5,C2:C5,D2:D5,E2:E5", false);
            stockChart.NSeries[0].Name = "Sample Stock";

            // Show data values on the chart
            stockChart.NSeries[0].DataLabels.ShowValue = true;

            // Save the workbook to a file
            string outputPath = "StockChartWithDataLabels.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
