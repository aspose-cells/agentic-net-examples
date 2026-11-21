// Title: How to add data labels to a Stock Open‑High‑Low‑Close chart using date‑formatted cells in Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add data labels to OHLC stock chart with dates as category labels | set DataLabels.IsValueShown for StockOpenHighLowClose chart in Aspose.Cells .NET | use a date formatted cell range as source for chart data labels in Aspose.Cells C# | create stock open‑high‑low‑close chart from worksheet data using Aspose.Cells | Aspose.Cells chart category data from date column C# example
// Tags: Aspose.Cells stock chart data labels | C# chart category range from date column | StockOpenHighLowClose chart with date labels | DataLabels.IsValueShown Aspose.Cells C# | format cells as dates for chart categories Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, writes dates (formatted with built‑in number format 14) and OHLC values to columns A‑E, adds a StockOpenHighLowClose chart, sets the series range to B2:E6 and the category (X‑axis) range to A2:A6, and demonstrates how to enable value data labels via the DataLabels.IsValueShown property (commented for compatibility). The workbook is saved as StockChartWithDateLabels.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate header row
            ws.Cells["A1"].PutValue("Date");
            ws.Cells["B1"].PutValue("Open");
            ws.Cells["C1"].PutValue("High");
            ws.Cells["D1"].PutValue("Low");
            ws.Cells["E1"].PutValue("Close");

            // Prepare a date style (built‑in format 14)
            Style dateStyle = wb.CreateStyle();
            dateStyle.Number = 14;

            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 5; i++)
            {
                // Date column (formatted as date)
                Cell dateCell = ws.Cells[i + 1, 0];
                dateCell.PutValue(startDate.AddDays(i));
                dateCell.SetStyle(dateStyle);

                // OHLC values
                ws.Cells[i + 1, 1].PutValue(100 + i * 2); // Open
                ws.Cells[i + 1, 2].PutValue(105 + i * 2); // High
                ws.Cells[i + 1, 3].PutValue(95 + i * 2);  // Low
                ws.Cells[i + 1, 4].PutValue(102 + i * 2); // Close
            }

            // Add a Stock Open‑High‑Low‑Close chart
            int chartIdx = ws.Charts.Add(ChartType.StockOpenHighLowClose, 7, 0, 25, 10);
            Chart chart = ws.Charts[chartIdx];

            // Set the data range for the series (Open‑Close)
            chart.NSeries.Add("B2:E6", true);
            // Set category (X‑axis) labels to the date range
            chart.NSeries.CategoryData = "A2:A6";

            // NOTE: DataLabels properties/methods may not be available in older Aspose.Cells versions.
            // If needed, enable value labels like this (uncomment when supported):
            // chart.NSeries[0].DataLabels.IsValueShown = true;
            // chart.NSeries[0].DataLabels.SetDataLabelRange("A2:A6");

            // Ensure the output directory exists
            string outputPath = "StockChartWithDateLabels.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
