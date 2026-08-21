// Title: Create a Candlestick Stock Chart with 20% Candle Width Using Aspose.Cells for .NET
// Description: Learn how to generate an Excel workbook with OHLC data, add a StockOpenHighLowClose (candlestick) chart, and set the candle width to 20 % of the bar spacing by configuring the GapWidth property in Aspose.Cells for C#.
// Keywords: Aspose.Cells candlestick chart C# | StockOpenHighLowClose GapWidth | set candle width Aspose.Cells | OHLC chart Aspose.Cells .NET | customize candlestick width Excel
// Common Searches: Aspose.Cells change stock chart to candlestick | set candlestick candle width 20 percent Aspose | GapWidth property candlestick chart .NET | create OHLC candlestick chart with Aspose.Cells | adjust candle width in Excel chart using C#
// Developer Intent: Create a candlestick (OHLC) chart and configure its candle width to 20 % of the default bar spacing with Aspose.Cells for .NET.
// Use Cases: Financial reporting: embed a candlestick chart that highlights price movements with a precise visual width. | Data visualization: fine‑tune candle thickness for better readability in stock analysis workbooks. | Automated Excel generation: produce Excel files with custom‑sized candlestick charts for distribution to investors or analysts.
// AI Prompts: Generate C# code that builds a candlestick chart from OHLC data using Aspose.Cells and sets the candle width to 20 % via GapWidth. | Explain how the GapWidth property controls candle width in a StockOpenHighLowClose chart and how to calculate the value for a desired percentage. | Show how to modify an existing Aspose.Cells chart to switch its type to candlestick and adjust the candle width without recreating the chart.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Learn how to generate an Excel workbook with OHLC data, add a StockOpenHighLowClose (candlestick) chart, and set the candle width to 20 % of the bar spacing by configuring the GapWidth property in Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data required for a stock (candlestick) chart
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Open");
            sheet.Cells["C1"].PutValue("High");
            sheet.Cells["D1"].PutValue("Low");
            sheet.Cells["E1"].PutValue("Close");

            sheet.Cells["A2"].PutValue("2023-01-01");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["C2"].PutValue(110);
            sheet.Cells["D2"].PutValue(95);
            sheet.Cells["E2"].PutValue(105);

            sheet.Cells["A3"].PutValue("2023-01-02");
            sheet.Cells["B3"].PutValue(105);
            sheet.Cells["C3"].PutValue(115);
            sheet.Cells["D3"].PutValue(100);
            sheet.Cells["E3"].PutValue(110);

            sheet.Cells["A4"].PutValue("2023-01-03");
            sheet.Cells["B4"].PutValue(110);
            sheet.Cells["C4"].PutValue(120);
            sheet.Cells["D4"].PutValue(108);
            sheet.Cells["E4"].PutValue(115);

            // Add a stock chart (Open‑High‑Low‑Close) which renders as a candlestick chart
            int chartIndex = sheet.Charts.Add(ChartType.StockOpenHighLowClose, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Define the series for Open, High, Low and Close
            int openSeriesIdx = chart.NSeries.Add("B2:B4", true);
            chart.NSeries[openSeriesIdx].Name = "Open";

            int highSeriesIdx = chart.NSeries.Add("C2:C4", true);
            chart.NSeries[highSeriesIdx].Name = "High";

            int lowSeriesIdx = chart.NSeries.Add("D2:D4", true);
            chart.NSeries[lowSeriesIdx].Name = "Low";

            int closeSeriesIdx = chart.NSeries.Add("E2:E4", true);
            chart.NSeries[closeSeriesIdx].Name = "Close";

            // Set the category (X‑axis) data – the dates
            chart.NSeries.CategoryData = "A2:A4";

            // Ensure the chart type is candlestick (same enum value)
            chart.Type = ChartType.StockOpenHighLowClose;

            // Adjust candle width: set GapWidth to 20 (20 % of the default bar width)
            chart.GapWidth = 20;

            // Determine output file path and ensure the directory exists
            string outputPath = "CandlestickChart.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook with the configured chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
