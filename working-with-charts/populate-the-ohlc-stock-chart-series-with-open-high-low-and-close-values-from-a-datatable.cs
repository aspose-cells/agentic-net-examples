// Title: Create an OHLC Stock Chart from a DataTable with Aspose.Cells for .NET
// Description: This example builds a DataTable containing Date, Open, High, Low, and Close values, writes it to the first worksheet, adds a StockOpenHighLowClose chart, links each series to the corresponding columns, sets the Date column as the category axis, enables high‑low lines, and saves the workbook as an Excel file.
// Keywords: Aspose.Cells | OHLC chart | StockOpenHighLowClose | DataTable to chart | C# | .NET | Excel financial chart | high‑low lines | category axis dates | chart series range
// Common Searches: Aspose.Cells bind DataTable to OHLC chart | Create StockOpenHighLowClose chart in C# | Set date axis for stock chart Aspose.Cells | Enable high low lines in Aspose.Cells chart | Populate OHLC series from worksheet data
// Developer Intent: Generate a Stock Open‑High‑Low‑Close chart in an Excel workbook using data stored in a DataTable.
// Use Cases: Transform market data retrieved from a database into a visual OHLC chart for financial reports. | Automate daily price‑movement charts for a securities‑analysis dashboard. | Batch‑process large sets of historical stock prices into Excel files with ready‑to‑use OHLC charts.
// AI Prompts: Write C# code that reads OHLC data from a CSV file into a DataTable and creates a StockOpenHighLowClose chart with Aspose.Cells. | Refactor the sample to use named ranges instead of hard‑coded cell references for chart series. | Show how to customize colors, markers, and line styles of an OHLC chart using the Aspose.Cells API.

using System;
using System.Data;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace OHLCChartExample
{
    // This example builds a DataTable containing Date, Open, High, Low, and Close values, writes it to the first worksheet, adds a StockOpenHighLowClose chart, links each series to the corresponding columns, sets the Date column as the category axis, enables high‑low lines, and saves the workbook as an Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Prepare a DataTable with OHLC data
                DataTable ohlcTable = new DataTable();
                ohlcTable.Columns.Add("Date", typeof(DateTime));
                ohlcTable.Columns.Add("Open", typeof(double));
                ohlcTable.Columns.Add("High", typeof(double));
                ohlcTable.Columns.Add("Low", typeof(double));
                ohlcTable.Columns.Add("Close", typeof(double));

                // Sample rows
                ohlcTable.Rows.Add(DateTime.Parse("2023-01-01"), 100.5, 105.2, 99.8, 104.0);
                ohlcTable.Rows.Add(DateTime.Parse("2023-01-02"), 104.0, 108.5, 103.2, 107.1);
                ohlcTable.Rows.Add(DateTime.Parse("2023-01-03"), 107.1, 110.0, 106.5, 109.3);
                ohlcTable.Rows.Add(DateTime.Parse("2023-01-04"), 109.3, 112.4, 108.7, 111.0);
                ohlcTable.Rows.Add(DateTime.Parse("2023-01-05"), 111.0, 115.2, 110.5, 114.8);

                // 2. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 3. Manually import the DataTable into the worksheet starting at cell A1
                int currentRow = 0;
                // Write column headers
                for (int col = 0; col < ohlcTable.Columns.Count; col++)
                {
                    sheet.Cells[currentRow, col].PutValue(ohlcTable.Columns[col].ColumnName);
                }
                currentRow++;

                // Write data rows
                foreach (DataRow dr in ohlcTable.Rows)
                {
                    sheet.Cells[currentRow, 0].PutValue((DateTime)dr[0]);
                    sheet.Cells[currentRow, 1].PutValue(Convert.ToDouble(dr[1]));
                    sheet.Cells[currentRow, 2].PutValue(Convert.ToDouble(dr[2]));
                    sheet.Cells[currentRow, 3].PutValue(Convert.ToDouble(dr[3]));
                    sheet.Cells[currentRow, 4].PutValue(Convert.ToDouble(dr[4]));
                    currentRow++;
                }

                // 4. Add a Stock Open‑High‑Low‑Close chart to the worksheet
                // Parameters: chart type, top row, left column, bottom row, right column
                int chartIndex = sheet.Charts.Add(ChartType.StockOpenHighLowClose, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // 5. Define the data ranges for each series (Open, High, Low, Close)
                chart.NSeries.Add("=Sheet1!$B$2:$B$6", true); // Open
                chart.NSeries.Add("=Sheet1!$C$2:$C$6", true); // High
                chart.NSeries.Add("=Sheet1!$D$2:$D$6", true); // Low
                chart.NSeries.Add("=Sheet1!$E$2:$E$6", true); // Close

                // 6. Set the category (X‑axis) data to the Date column
                chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$6";

                // 7. Configure each series as a Stock Open‑High‑Low‑Close series and enable high‑low lines
                foreach (Series s in chart.NSeries)
                {
                    s.Type = ChartType.StockOpenHighLowClose;
                    s.HasHiLoLines = true;
                }

                // 8. Save the workbook
                workbook.Save("OHLC_StockChart.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
