using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet to store OHLC data
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Header row
            dataSheet.Cells["A1"].PutValue("Date");
            dataSheet.Cells["B1"].PutValue("Open");
            dataSheet.Cells["C1"].PutValue("High");
            dataSheet.Cells["D1"].PutValue("Low");
            dataSheet.Cells["E1"].PutValue("Close");

            // Sample data rows (Day 1 to Day 5)
            for (int i = 2; i <= 6; i++)
            {
                dataSheet.Cells[$"A{i}"].PutValue($"Day {i - 1}");
                dataSheet.Cells[$"B{i}"].PutValue(100 + i * 2); // Open
                dataSheet.Cells[$"C{i}"].PutValue(110 + i * 2); // High
                dataSheet.Cells[$"D{i}"].PutValue(90 + i * 2);  // Low
                dataSheet.Cells[$"E{i}"].PutValue(105 + i * 2); // Close
            }

            // Add a new chart sheet (type Chart)
            int chartSheetIndex = workbook.Worksheets.Add(SheetType.Chart);
            Worksheet chartSheet = workbook.Worksheets[chartSheetIndex];
            chartSheet.Name = "OHLCChart";

            // Add an OHLC (Open‑High‑Low‑Close) stock chart to the chart sheet
            int chartIndex = chartSheet.Charts.Add(ChartType.StockOpenHighLowClose, 0, 0, 30, 15);
            Chart chart = chartSheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("Data!$B$2:$E$6", true);          // Open, High, Low, Close
            chart.NSeries.CategoryData = "Data!$A$2:$A$6";     // Dates

            // Optional: give the chart a title
            chart.Title.Text = "OHLC Stock Chart";

            // Save the workbook (ensure the directory exists)
            string outputPath = "OHLCChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}