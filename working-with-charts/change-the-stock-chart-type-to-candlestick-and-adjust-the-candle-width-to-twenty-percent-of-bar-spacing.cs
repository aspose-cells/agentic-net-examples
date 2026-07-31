// Title: Aspose.Cells for .NET – Build a Candlestick (OHLC) Stock Chart with 20 % Candle Width
// Description: This .NET example shows how to create a new workbook, populate it with date and OHLC values, add a StockOpenHighLowClose chart, bind the series and categories, set the GapWidth to 20 (≈20 % of bar spacing), give the chart a title, and save the file as CandlestickChart.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells candlestick chart .NET | StockOpenHighLowClose Aspose.Cells | set GapWidth Aspose.Cells | OHLC chart example C# | financial chart Excel Aspose
// Common Searches: how to create a candlestick chart with Aspose.Cells | set candle width percentage in Aspose.Cells .NET | Aspose.Cells StockOpenHighLowClose example | adjust GapWidth for stock charts in C# | generate OHLC Excel chart programmatically
// Developer Intent: Programmatically generate an Excel workbook that displays OHLC data as a candlestick chart with a custom 20 % candle width.
// Use Cases: Automate financial reporting by visualizing daily open‑high‑low‑close data with thin candlesticks for clearer trends. | Create a reusable Excel template that adds a candlestick chart and controls candle thickness via GapWidth. | Export trading system outputs to Excel where the chart title and spacing are pre‑configured for presentations.
// AI Prompts: Show how to switch the chart to StockHighLowClose while keeping the 20 % candle width. | Provide code that reads OHLC rows from a CSV file and builds a candlestick chart with GapWidth set to 15. | Explain the relationship between GapWidth and candle width and how to compute the value for a desired percentage.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This .NET example shows how to create a new workbook, populate it with date and OHLC values, add a StockOpenHighLowClose chart, bind the series and categories, set the GapWidth to 20 (≈20 % of bar spacing), give the chart a title, and save the file as CandlestickChart.xlsx using Aspose.Cells.
    public class StockCandlestickChartDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a stock chart (Date, Open, High, Low, Close)
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Open");
                sheet.Cells["C1"].PutValue("High");
                sheet.Cells["D1"].PutValue("Low");
                sheet.Cells["E1"].PutValue("Close");

                // Example rows
                sheet.Cells["A2"].PutValue(DateTime.Today.AddDays(-4));
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["C2"].PutValue(110);
                sheet.Cells["D2"].PutValue(95);
                sheet.Cells["E2"].PutValue(105);

                sheet.Cells["A3"].PutValue(DateTime.Today.AddDays(-3));
                sheet.Cells["B3"].PutValue(105);
                sheet.Cells["C3"].PutValue(115);
                sheet.Cells["D3"].PutValue(100);
                sheet.Cells["E3"].PutValue(110);

                sheet.Cells["A4"].PutValue(DateTime.Today.AddDays(-2));
                sheet.Cells["B4"].PutValue(110);
                sheet.Cells["C4"].PutValue(120);
                sheet.Cells["D4"].PutValue(108);
                sheet.Cells["E4"].PutValue(115);

                sheet.Cells["A5"].PutValue(DateTime.Today.AddDays(-1));
                sheet.Cells["B5"].PutValue(115);
                sheet.Cells["C5"].PutValue(125);
                sheet.Cells["D5"].PutValue(112);
                sheet.Cells["E5"].PutValue(118);

                // Add a candlestick (Open‑High‑Low‑Close) stock chart
                int chartIndex = sheet.Charts.Add(ChartType.StockOpenHighLowClose, 6, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart (Open, High, Low, Close)
                chart.NSeries.Add("B2:E5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Adjust candle width to 20% of the bar spacing
                chart.GapWidth = 20;

                // Optional: give the chart a title
                chart.Title.Text = "Candlestick Chart Example";

                // Save the workbook
                string outputPath = "CandlestickChart.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while creating the candlestick chart: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            StockCandlestickChartDemo.Run();
        }
    }
}
