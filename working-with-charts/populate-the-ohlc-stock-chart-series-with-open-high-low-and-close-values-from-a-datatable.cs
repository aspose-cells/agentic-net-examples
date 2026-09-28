// Title: How to fill an Aspose.Cells StockOpenHighLowClose chart with open, high, low, and close data from a DataTable in C#
// AI Prompts: Write C# code that writes a DataTable containing Date, Open, High, Low, and Close columns to a worksheet and binds each column to the corresponding series of an Aspose.Cells StockOpenHighLowClose chart. | Show how to calculate dynamic range strings for chart series based on the number of rows in a DataTable and assign them to NSeries in Aspose.Cells. | Demonstrate setting the category (X‑axis) data to the Date column and naming the OHLC series programmatically using Aspose.Cells for .NET.
// Common Searches: aspnet populate OHLC chart series from DataTable using Aspose.Cells | c# Aspose.Cells StockOpenHighLowClose chart bind to DataTable rows | dynamic range address for Aspose.Cells chart series based on DataTable count | set category data for Aspose.Cells stock chart from date column | save workbook with OHLC chart Aspose.Cells example
// Tags: Aspose.Cells StockOpenHighLowClose series binding | C# dynamic chart range calculation Aspose.Cells | assign category data Aspose.Cells stock chart | write DataTable to worksheet for charting Aspose.Cells | set series names programmatically Aspose.Cells

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, writes stock data from a DataTable to the first worksheet, builds range strings for Date, Open, High, Low, and Close columns, adds a StockOpenHighLowClose chart, sets the Date range as category data, adds the four OHLC series with the calculated ranges, assigns series names, and saves the file as OHLC_StockChart.xlsx.
class OhlcChartExample
{
    static void Main()
    {
        try
        {
            // Obtain stock data.
            DataTable stockTable = GetStockDataTable();

            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Write headers.
            for (int col = 0; col < stockTable.Columns.Count; col++)
            {
                sheet.Cells[0, col].PutValue(stockTable.Columns[col].ColumnName);
            }

            // Write data rows.
            for (int row = 0; row < stockTable.Rows.Count; row++)
            {
                for (int col = 0; col < stockTable.Columns.Count; col++)
                {
                    sheet.Cells[row + 1, col].PutValue(stockTable.Rows[row][col]);
                }
            }

            // Determine the number of data rows (excluding header).
            int dataRowCount = stockTable.Rows.Count;

            // Define the range addresses for each column (including header row).
            string dateRange  = $"A2:A{dataRowCount + 1}";
            string openRange  = $"B2:B{dataRowCount + 1}";
            string highRange  = $"C2:C{dataRowCount + 1}";
            string lowRange   = $"D2:D{dataRowCount + 1}";
            string closeRange = $"E2:E{dataRowCount + 1}";

            // Add an OHLC stock chart to the worksheet.
            // Position: from row 5, column 0 to row 20, column 10.
            int chartIndex = sheet.Charts.Add(ChartType.StockOpenHighLowClose, 5, 0, 20, 10);
            Chart ohlcChart = sheet.Charts[chartIndex];

            // Set the X‑axis (category) data – dates.
            ohlcChart.NSeries.CategoryData = dateRange;

            // Add series for Open, High, Low, and Close values.
            ohlcChart.NSeries.Add(openRange, true);
            ohlcChart.NSeries.Add(highRange, true);
            ohlcChart.NSeries.Add(lowRange, true);
            ohlcChart.NSeries.Add(closeRange, true);

            // Optionally set series names (taken from the header row).
            ohlcChart.NSeries[0].Name = "Open";
            ohlcChart.NSeries[1].Name = "High";
            ohlcChart.NSeries[2].Name = "Low";
            ohlcChart.NSeries[3].Name = "Close";

            // Save the workbook to a file.
            string outputPath = "OHLC_StockChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Mock method to obtain a DataTable with sample stock data.
    // Replace this with actual data retrieval logic.
    static DataTable GetStockDataTable()
    {
        DataTable dt = new DataTable();
        dt.Columns.Add("Date", typeof(DateTime));
        dt.Columns.Add("Open", typeof(double));
        dt.Columns.Add("High", typeof(double));
        dt.Columns.Add("Low", typeof(double));
        dt.Columns.Add("Close", typeof(double));

        // Sample rows
        dt.Rows.Add(new DateTime(2023, 1, 2), 150.0, 155.0, 149.0, 154.0);
        dt.Rows.Add(new DateTime(2023, 1, 3), 154.0, 158.0, 152.0, 157.0);
        dt.Rows.Add(new DateTime(2023, 1, 4), 157.0, 160.0, 156.0, 159.0);
        // Add more rows as needed...

        return dt;
    }
}
