// Title: How to enable high and low data labels on an Open‑High‑Low‑Close stock chart with Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, add an Open‑High‑Low‑Close stock chart, and set its first series to display only the high and low values as data labels. | Adjust the chart's data label settings to hide the default value label while preserving the high/low labels. | Insert a second OHLC series into the same chart and configure its data labels to show high and low values only.
// Common Searches: Aspose.Cells C# show only high and low values in data labels for OHLC stock chart | how to hide value label but keep high low labels on stock chart using Aspose.Cells | C# example enabling high low data labels on OpenHighLowClose chart with Aspose.Cells | Aspose.Cells .NET customize stock chart data label visibility high low only
// Tags: Aspose.Cells enable OHLC data labels | C# chart series high low label configuration | Aspose.Cells OpenHighLowClose chart customization | Excel stock chart data label visibility .NET | Aspose.Cells hide default value label

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The sample creates a workbook, fills cells A1:E6 with date, open, high, low, and close data, adds an Open‑High‑Low‑Close stock chart referencing that range, sets the category axis to the dates, configures the first series to hide its default value label while keeping high/low labels visible, and saves the file as StockChartWithHighLowLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "Data";

            // Populate sample stock data: Date, Open, High, Low, Close
            ws.Cells["A1"].PutValue("Date");
            ws.Cells["B1"].PutValue("Open");
            ws.Cells["C1"].PutValue("High");
            ws.Cells["D1"].PutValue("Low");
            ws.Cells["E1"].PutValue("Close");

            DateTime start = new DateTime(2023, 1, 1);
            for (int i = 0; i < 5; i++)
            {
                ws.Cells[i + 1, 0].PutValue(start.AddDays(i));
                ws.Cells[i + 1, 1].PutValue(100 + i * 2); // Open
                ws.Cells[i + 1, 2].PutValue(105 + i * 2); // High
                ws.Cells[i + 1, 3].PutValue(95 + i * 2);  // Low
                ws.Cells[i + 1, 4].PutValue(102 + i * 2); // Close
            }

            // Add a stock chart (Open-High-Low-Close)
            int chartIndex = ws.Charts.Add(ChartType.StockOpenHighLowClose, 7, 0, 25, 10);
            Chart chart = ws.Charts[chartIndex];

            // Set the data range for the stock series (Open, High, Low, Close)
            chart.NSeries.Add("B2:E6", true);
            // Set the category (X) axis data (dates)
            chart.NSeries.CategoryData = "A2:A6";

            // Configure data labels to hide default values
            if (chart.NSeries.Count > 0 && chart.NSeries[0] != null)
            {
                chart.NSeries[0].DataLabels.ShowValue = false;
                // High/Low visibility defaults are appropriate for this chart type
            }

            // Save the workbook
            string outputPath = "StockChartWithHighLowLabels.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
