// Title: Aspose.Cells for .NET – Add High & Low Data Labels to a Stock Open‑High‑Low‑Close Chart (C#)
// Description: Creates a new workbook, fills it with dates and OHLC prices, inserts a StockOpenHighLowClose chart on a chart sheet, adds Open, High, Low and Close series, sets the date axis, and enables data labels that show the High values above the points and the Low values below them before saving the file.
// Keywords: Aspose.Cells C# stock chart | Excel data labels high low | StockOpenHighLowClose Aspose | show high values chart | show low values chart | label position above below | financial chart automation | OHLC chart Aspose.Cells
// Common Searches: Aspose.Cells add data labels to stock chart | display high and low values on Excel stock chart .NET | set label position for high series Aspose.Cells | enable low series data labels in StockOpenHighLowClose chart | C# code for OHLC chart with data labels
// Developer Intent: Add data labels that display high and low values on specific series of a stock chart.
// Use Cases: Generate daily financial reports where extreme price points are highlighted directly on the chart. | Automate Excel dashboards that need clear high/low markers for quick market analysis. | Create investment presentations with stock charts that label price highs above and lows below each data point.
// AI Prompts: Write C# using Aspose.Cells to add data labels to the High and Low series of a StockOpenHighLowClose chart and set their positions above and below the points. | Show how to save the workbook to a MemoryStream while keeping high and low data labels enabled. | Demonstrate customizing font size, color, and style of high and low data labels on a stock chart with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a new workbook, fills it with dates and OHLC prices, inserts a StockOpenHighLowClose chart on a chart sheet, adds Open, High, Low and Close series, sets the date axis, and enables data labels that show the High values above the points and the Low values below them before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate headers
                sheet.Cells["A1"].PutValue("Date");
                sheet.Cells["B1"].PutValue("Open");
                sheet.Cells["C1"].PutValue("High");
                sheet.Cells["D1"].PutValue("Low");
                sheet.Cells["E1"].PutValue("Close");

                // Sample dates
                sheet.Cells["A2"].PutValue("2023-01-01");
                sheet.Cells["A3"].PutValue("2023-01-02");
                sheet.Cells["A4"].PutValue("2023-01-03");
                sheet.Cells["A5"].PutValue("2023-01-04");

                // Sample price data
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["B3"].PutValue(102);
                sheet.Cells["B4"].PutValue(101);
                sheet.Cells["B5"].PutValue(103);

                sheet.Cells["C2"].PutValue(105);
                sheet.Cells["C3"].PutValue(108);
                sheet.Cells["C4"].PutValue(107);
                sheet.Cells["C5"].PutValue(110);

                sheet.Cells["D2"].PutValue(95);
                sheet.Cells["D3"].PutValue(98);
                sheet.Cells["D4"].PutValue(97);
                sheet.Cells["D5"].PutValue(99);

                sheet.Cells["E2"].PutValue(102);
                sheet.Cells["E3"].PutValue(106);
                sheet.Cells["E4"].PutValue(103);
                sheet.Cells["E5"].PutValue(108);

                // Add a chart sheet (type Chart) for the stock chart
                int chartSheetIndex = workbook.Worksheets.Add(SheetType.Chart);
                Worksheet chartSheet = workbook.Worksheets[chartSheetIndex];

                // Create the stock chart on the chart sheet
                int chartIndex = chartSheet.Charts.Add(ChartType.StockOpenHighLowClose, 5, 0, 25, 15);
                Chart stockChart = chartSheet.Charts[chartIndex];

                // Add series for Open, High, Low, Close
                stockChart.NSeries.Add("B2:B5", true); // Open
                stockChart.NSeries.Add("C2:C5", true); // High
                stockChart.NSeries.Add("D2:D5", true); // Low
                stockChart.NSeries.Add("E2:E5", true); // Close

                // Set category (X) axis data (dates)
                stockChart.NSeries.CategoryData = "A2:A5";

                // Enable data labels for the High and Low series
                Series highSeries = stockChart.NSeries[1]; // High
                Series lowSeries = stockChart.NSeries[2];  // Low

                highSeries.DataLabels.ShowValue = true;
                lowSeries.DataLabels.ShowValue = true;

                // Position the data labels for better readability
                highSeries.DataLabels.Position = LabelPositionType.Above;
                lowSeries.DataLabels.Position = LabelPositionType.Below;

                // Define output file path
                string outputPath = "StockChartWithHighLowDataLabels.xlsx";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (string.IsNullOrEmpty(outputDir))
                {
                    outputDir = Directory.GetCurrentDirectory();
                }
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the stock chart:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
