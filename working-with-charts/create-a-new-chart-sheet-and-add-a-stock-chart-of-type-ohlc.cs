// Title: Generate a separate chart sheet with an OHLC (Open‑High‑Low‑Close) stock chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to add a new worksheet as a chart sheet and place a StockOpenHighLowClose chart that pulls its data from another sheet. | Show how to configure the category axis to a date range and assign custom series names (Open, High, Low, Close) for an OHLC chart on a dedicated chart sheet with Aspose.Cells.
// Common Searches: how to add an OHLC stock chart on a separate worksheet with Aspose.Cells C# | Aspose.Cells create chart sheet for Open High Low Close data | C# example of StockOpenHighLowClose chart on its own sheet using Aspose.Cells | set category axis dates for OHLC chart in Aspose.Cells workbook | save workbook with chart sheet containing OHLC chart Aspose.Cells .NET
// Tags: Aspose.Cells separate chart worksheet | StockOpenHighLowClose chart generation C# | OHLC chart data source Aspose.Cells | date category axis Aspose.Cells | custom series names OHLC Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsStockChartExample
{
    // The example creates a workbook, fills a data sheet with Date, Open, High, Low, and Close columns, adds a new worksheet that serves as a chart sheet, inserts a StockOpenHighLowClose (OHLC) chart referencing the data ranges, sets the chart title and date-based category axis, defines four series with names Open, High, Low, and Close, and saves the file as OHLC_Chart_Sheet.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the default worksheet (Sheet1) to store sample data
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Populate sample data for OHLC chart
                dataSheet.Cells["A1"].PutValue("Date");
                dataSheet.Cells["B1"].PutValue("Open");
                dataSheet.Cells["C1"].PutValue("High");
                dataSheet.Cells["D1"].PutValue("Low");
                dataSheet.Cells["E1"].PutValue("Close");

                dataSheet.Cells["A2"].PutValue("2023-01-01");
                dataSheet.Cells["A3"].PutValue("2023-01-02");
                dataSheet.Cells["A4"].PutValue("2023-01-03");
                dataSheet.Cells["A5"].PutValue("2023-01-04");

                dataSheet.Cells["B2"].PutValue(120);
                dataSheet.Cells["B3"].PutValue(125);
                dataSheet.Cells["B4"].PutValue(123);
                dataSheet.Cells["B5"].PutValue(128);

                dataSheet.Cells["C2"].PutValue(130);
                dataSheet.Cells["C3"].PutValue(135);
                dataSheet.Cells["C4"].PutValue(133);
                dataSheet.Cells["C5"].PutValue(138);

                dataSheet.Cells["D2"].PutValue(115);
                dataSheet.Cells["D3"].PutValue(118);
                dataSheet.Cells["D4"].PutValue(119);
                dataSheet.Cells["D5"].PutValue(122);

                dataSheet.Cells["E2"].PutValue(128);
                dataSheet.Cells["E3"].PutValue(132);
                dataSheet.Cells["E4"].PutValue(130);
                dataSheet.Cells["E5"].PutValue(135);

                // Add a new worksheet that will contain the chart
                Worksheet chartSheet = workbook.Worksheets.Add("OHLCChart");

                // Add a Stock Open‑High‑Low‑Close chart (OHLC)
                // The Charts.Add method expects an int for the chart type, so cast the enum.
                int chartIndex = chartSheet.Charts.Add(0, 0, 15, 10, (int)ChartType.StockOpenHighLowClose);
                Chart ohlcChart = chartSheet.Charts[chartIndex];

                // Set the chart title
                ohlcChart.Title.Text = "Sample OHLC Stock Chart";

                // Set the categories (X‑axis) to the dates column
                ohlcChart.NSeries.CategoryData = "Data!A2:A5";

                // Add series for Open, High, Low, and Close values
                ohlcChart.NSeries.Add("Data!B2:B5", false); // Open
                ohlcChart.NSeries.Add("Data!C2:C5", false); // High
                ohlcChart.NSeries.Add("Data!D2:D5", false); // Low
                ohlcChart.NSeries.Add("Data!E2:E5", false); // Close

                // Set series names (displayed in the legend)
                ohlcChart.NSeries[0].Name = "Open";
                ohlcChart.NSeries[1].Name = "High";
                ohlcChart.NSeries[2].Name = "Low";
                ohlcChart.NSeries[3].Name = "Close";

                // Define output file path
                string outputPath = "OHLC_Chart_Sheet.xlsx";

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
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the OHLC chart workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
